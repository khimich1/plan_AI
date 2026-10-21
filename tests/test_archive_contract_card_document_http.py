"""HTTP блока договора по contract_id: реквизиты, разбор, Word, PDF и ЭДО (CDC-010)."""

from __future__ import annotations

import inspect
from datetime import date
from io import BytesIO
from pathlib import Path

import fastapi.dependencies.utils as _fastapi_dependencies
import pytest
from openpyxl import Workbook

def _get_typed_signature(call: object) -> inspect.Signature:
    """Класс-зависимость не имеет __globals__, отложенный UploadFile иначе не резолвится."""
    signature = inspect.signature(call)  # type: ignore[arg-type]
    globalns = getattr(call, "__globals__", None) or {}
    if not globalns and isinstance(call, type):
        init = getattr(call, "__init__", None)
        globalns = getattr(init, "__globals__", {}) or {}
    typed_params = [
        inspect.Parameter(
            name=param.name,
            kind=param.kind,
            default=param.default,
            annotation=_fastapi_dependencies.get_typed_annotation(param.annotation, globalns),
        )
        for param in signature.parameters.values()
    ]
    return inspect.Signature(typed_params)


_fastapi_dependencies.get_typed_signature = _get_typed_signature  # type: ignore[assignment]

from app.core.settings import get_settings
from core.kp_db_schema import ensure_schema
from app.dependencies.services import get_supply_contract_service
from app.main import create_app
from app.security.session import create_session_token
from app.services.supply_contract_service import SupplyContractService
from core.kp_db_common import _connect
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient

_CARDS = "/api/v1/commercial/archive/contract-cards"
_CONTRACTS = "/api/v1/commercial/archive/supply-contracts"
_FILL_REQUISITES = "Сначала заполните реквизиты договора"
_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

_USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
        "session_version": 0,
    },
    {
        "id": 2,
        "username": "manager",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
        "session_version": 0,
    },
    {
        "id": 3,
        "username": "shop",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
        "session_version": 0,
    },
]

_BLANK_COLUMNS = (
    "legal_form",
    "full_name",
    "short_name",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "inn",
    "ogrn",
    "legal_address",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
    "contract_date",
)


def _cookie(user_id: int, username: str, role: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }


def _buyer(**overrides: object) -> dict[str, object]:
    data: dict[str, object] = {
        "legal_form": "ooo",
        "full_name": "ООО Ромашка",
        "short_name": "Ромашка",
        "signatory_position": "Генерального директора",
        "signatory_name": "Иванов Иван",
        "signatory_verb": "действующего",
        "authority_basis": "устав",
        "inn": "7701000001",
        "kpp": "770101001",
        "ogrn": "1027700000001",
        "legal_address": "Москва",
        "email": "a@b.ru",
        "bank_name": "Банк",
        "account": "40702810000000000001",
        "corr_account": "30101810000000000001",
        "bik": "044525225",
    }
    data.update(overrides)
    return data


def _count(db_path: str, table: str) -> int:
    with _connect(db_path) as conn:
        row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    return int(row[0])


def _set_columns(db_path: str, contract_id: int, **columns: object) -> None:
    assignments = ", ".join(f"{name} = ?" for name in columns)
    with _connect(db_path) as conn:
        conn.execute(
            f"UPDATE supply_contract SET {assignments} WHERE id = ?",
            (*columns.values(), int(contract_id)),
        )


def _card_xlsx() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "ИНН 7604010011"
    sheet["A2"] = "Р/с 40702810000000000007"
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _forbid_libreoffice(monkeypatch: pytest.MonkeyPatch) -> None:
    def fail_run(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("soffice must not run")

    monkeypatch.setattr(
        "app.services.supply_contract_service.shutil.which",
        lambda name: "/usr/bin/soffice" if name == "soffice" else None,
    )
    monkeypatch.setattr("app.services.supply_contract_service.subprocess.run", fail_run)


@pytest.fixture()
def cards(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[CsrfAwareTestClient, str]:
    db_path = str(tmp_path / "cards.db")
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    assert get_settings().plita_db_path.resolve() == Path(db_path).resolve()
    patch_auth_users(monkeypatch, _USERS)
    app = create_app()
    app.dependency_overrides[get_supply_contract_service] = lambda: SupplyContractService(
        db_path=db_path,
        today=date(2026, 10, 20),
    )
    return CsrfAwareTestClient(app), db_path


def _create_card(
    client: CsrfAwareTestClient,
    cookies: dict[str, str],
    *,
    name: str = "ИП Секлюцкий",
    number: str = "0786/1/25",
) -> dict:
    created = client.post(
        _CARDS,
        json={"name": name, "number": number},
        cookies=cookies,
    )
    assert created.status_code == 200
    return created.json()


def test_contract_block_routes_keep_import_above_card_id_and_kp_routes() -> None:
    from app.api.v1.endpoints import archive as archive_routes

    paths = [route.path for route in archive_routes._archive_router.routes]
    import_path = "/commercial/archive/contract-cards/import"
    parse_path = "/commercial/archive/contract-cards/parse"
    card_path = "/commercial/archive/contract-cards/{card_id}"
    assert import_path in paths
    assert parse_path in paths
    assert paths.index(import_path) < paths.index(card_path)
    assert paths.index(parse_path) < paths.index(card_path)
    for suffix in ("requisites", "parse", "document", "edo-agreement"):
        assert f"/commercial/archive/supply-contracts/{{contract_id}}/{suffix}" in paths
    assert "/commercial/archive/{kp_id}/supply-contract" in paths
    assert "/commercial/archive/{kp_id}/supply-contract/document" in paths
    assert "/commercial/archive/{kp_id}/supply-contract/edo-agreement" in paths
    assert "/commercial/archive/{kp_id}/supply-contract/parse" in paths


def test_save_requisites_returns_the_same_card_with_the_previous_number(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    manager = _cookie(2, "manager", "manager")
    created = _create_card(client, manager)
    contract_id = created["id"]
    parties_before = _count(db_path, "counterparties")

    saved = client.post(
        f"{_CONTRACTS}/{contract_id}/requisites",
        json=_buyer(),
        cookies=manager,
    )
    again = client.post(
        f"{_CONTRACTS}/{contract_id}/requisites",
        json=_buyer(full_name="ООО Ромашка обновлённая"),
        cookies=manager,
    )

    assert saved.status_code == 200
    assert again.status_code == 200
    body = saved.json()
    updated = again.json()
    assert body["id"] == updated["id"] == contract_id
    assert body["number"] == updated["number"] == "0786/1/25"
    assert body["contract_date"] == "2025-01-01"
    assert updated["contract_date"] == "2025-01-01"
    assert body["legal_form"] == "ooo"
    assert body["full_name"] == "ООО Ромашка"
    assert updated["full_name"] == "ООО Ромашка обновлённая"
    assert body["inn"] == "7701000001"
    assert body["email"] == "a@b.ru"
    assert body["manager_name"] == "manager"
    assert updated["manager_name"] == "manager"
    assert body["counterparty_id"] is None
    assert body["status"] == "нет"
    assert "scan_path" not in body
    assert _count(db_path, "supply_contract") == 1
    assert _count(db_path, "counterparties") == parties_before

    card = client.get(f"{_CARDS}/{contract_id}", cookies=manager)
    assert card.status_code == 200
    assert card.json()["name"] == "ИП Секлюцкий"
    assert card.json()["number"] == "0786/1/25"
    assert card.json()["full_name"] == "ООО Ромашка обновлённая"
    assert card.json()["contract_date"] == "2025-01-01"


def test_save_requisites_changes_the_date_only_when_the_body_has_one(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = _create_card(client, admin)

    response = client.post(
        f"{_CONTRACTS}/{created['id']}/requisites",
        json=_buyer(contract_date="2026-04-15"),
        cookies=admin,
    )

    assert response.status_code == 200
    body = response.json()
    assert body["id"] == created["id"]
    assert body["number"] == "0786/1/25"
    assert body["contract_date"] == "2026-04-15"
    assert body["manager_name"] == "admin"
    assert _count(db_path, "supply_contract") == 1


def test_invalid_requisites_are_422_and_leave_the_card_blank(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = _create_card(client, admin)

    response = client.post(
        f"{_CONTRACTS}/{created['id']}/requisites",
        json=_buyer(kpp=None),
        cookies=admin,
    )

    assert response.status_code == 422
    assert "КПП" in response.json()["detail"]
    card = client.get(f"{_CARDS}/{created['id']}", cookies=admin)
    assert card.status_code == 200
    assert card.json()["legal_form"] is None
    assert card.json()["manager_name"] is None
    assert card.json()["number"] == "0786/1/25"
    assert _count(db_path, "supply_contract") == 1


@pytest.mark.parametrize("file_format", ["docx", "pdf"])
@pytest.mark.parametrize("kind", ["document", "edo-agreement"])
def test_download_before_requisites_is_400_and_does_not_call_libreoffice(
    cards: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
    file_format: str,
    kind: str,
) -> None:
    client, _db_path = cards
    _forbid_libreoffice(monkeypatch)
    admin = _cookie(1, "admin", "admin")
    created = _create_card(client, admin)

    response = client.get(
        f"{_CONTRACTS}/{created['id']}/{kind}",
        params={"format": file_format},
        cookies=admin,
    )

    assert response.status_code == 400
    assert response.json()["detail"] == _FILL_REQUISITES


def test_whitespace_requisites_download_is_400_without_libreoffice(
    cards: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, db_path = cards
    _forbid_libreoffice(monkeypatch)
    admin = _cookie(1, "admin", "admin")
    created = _create_card(client, admin)
    _set_columns(
        db_path,
        int(created["id"]),
        **{name: "   " for name in _BLANK_COLUMNS},
    )

    document = client.get(
        f"{_CONTRACTS}/{created['id']}/document",
        params={"format": "pdf"},
        cookies=admin,
    )
    edo = client.get(
        f"{_CONTRACTS}/{created['id']}/edo-agreement",
        params={"format": "docx"},
        cookies=admin,
    )

    assert document.status_code == 400
    assert document.json()["detail"] == _FILL_REQUISITES
    assert edo.status_code == 400
    assert edo.json()["detail"] == _FILL_REQUISITES


def test_docx_downloads_after_requisites_return_word_files(
    cards: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, _db_path = cards
    _forbid_libreoffice(monkeypatch)
    manager = _cookie(2, "manager", "manager")
    created = _create_card(client, manager)
    saved = client.post(
        f"{_CONTRACTS}/{created['id']}/requisites",
        json=_buyer(),
        cookies=manager,
    )
    assert saved.status_code == 200

    document = client.get(
        f"{_CONTRACTS}/{created['id']}/document",
        cookies=manager,
    )
    edo = client.get(
        f"{_CONTRACTS}/{created['id']}/edo-agreement",
        cookies=manager,
    )

    assert document.status_code == 200
    assert document.content.startswith(b"PK")
    assert document.headers["content-type"].startswith(_DOCX)
    assert 'filename="dogovor-0786-1-25.docx"' in document.headers["content-disposition"]
    assert edo.status_code == 200
    assert edo.content.startswith(b"PK")
    assert edo.headers["content-type"].startswith(_DOCX)
    assert 'filename="soglashenie-edo-0786-1-25.docx"' in edo.headers["content-disposition"]


def test_parse_card_file_does_not_need_kp_or_counterparty(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    manager = _cookie(2, "manager", "manager")
    created = _create_card(client, manager)

    response = client.post(
        f"{_CONTRACTS}/{created['id']}/parse",
        files={"file": ("card.xlsx", _card_xlsx(), _XLSX)},
        cookies=manager,
    )

    assert response.status_code == 200
    body = response.json()
    assert body["fields"]["inn"] == "7604010011"
    assert body["fields"]["account"] == "40702810000000000007"
    card = client.get(f"{_CARDS}/{created['id']}", cookies=manager)
    assert card.status_code == 200
    assert card.json()["legal_form"] is None
    assert card.json()["counterparty_id"] is None
    assert card.json()["number"] == "0786/1/25"
    assert _count(db_path, "supply_contract") == 1
    assert _count(db_path, "counterparties") == 0
    with _connect(db_path) as conn:
        offers = conn.execute("SELECT COUNT(*) FROM KP_offers").fetchone()
    assert int(offers[0]) == 0


def test_empty_parse_file_is_400(cards: tuple[CsrfAwareTestClient, str]) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = _create_card(client, admin)

    response = client.post(
        f"{_CONTRACTS}/{created['id']}/parse",
        files={"file": ("card.xlsx", b"", _XLSX)},
        cookies=admin,
    )

    assert response.status_code == 400
    assert response.json()["detail"] == "Пустой файл."


def test_missing_contract_block_routes_are_404(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")

    saved = client.post(f"{_CONTRACTS}/404/requisites", json=_buyer(), cookies=admin)
    document = client.get(f"{_CONTRACTS}/404/document", cookies=admin)
    parsed = client.post(
        f"{_CONTRACTS}/404/parse",
        files={"file": ("card.xlsx", _card_xlsx(), _XLSX)},
        cookies=admin,
    )

    assert saved.status_code == 404
    assert saved.json()["detail"] == "Договор №404 не найден"
    assert document.status_code == 404
    assert document.json()["detail"] == "Договор №404 не найден"
    assert parsed.status_code == 404
    assert parsed.json()["detail"] == "Договор №404 не найден"


def test_anonymous_contract_block_routes_are_401(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards

    assert client.post(f"{_CONTRACTS}/1/requisites", json=_buyer()).status_code == 401
    assert client.get(f"{_CONTRACTS}/1/document").status_code == 401
    assert client.get(f"{_CONTRACTS}/1/edo-agreement").status_code == 401
    assert client.post(
        f"{_CONTRACTS}/1/parse",
        files={"file": ("card.xlsx", b"x", _XLSX)},
    ).status_code == 401


def test_production_role_cannot_use_contract_block_routes(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    shop = _cookie(3, "shop", "production")

    assert client.post(
        f"{_CONTRACTS}/1/requisites",
        json=_buyer(),
        cookies=shop,
    ).status_code == 403
    assert client.get(f"{_CONTRACTS}/1/document", cookies=shop).status_code == 403
    assert client.get(f"{_CONTRACTS}/1/edo-agreement", cookies=shop).status_code == 403
    assert client.post(
        f"{_CONTRACTS}/1/parse",
        files={"file": ("card.xlsx", b"x", _XLSX)},
        cookies=shop,
    ).status_code == 403


def test_unknown_document_format_is_422(
    cards: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, _db_path = cards
    _forbid_libreoffice(monkeypatch)
    admin = _cookie(1, "admin", "admin")
    created = _create_card(client, admin)

    response = client.get(
        f"{_CONTRACTS}/{created['id']}/document",
        params={"format": "xls"},
        cookies=admin,
    )

    assert response.status_code == 422


def _passport_xlsx() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "Паспорт гражданина"
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


async def _block_card_vision(payload: bytes, *, filename: str) -> dict:
    raise AssertionError(f"картинка не должна распознаваться: {filename} {len(payload)}")


def _parse_service(db_path: str) -> SupplyContractService:
    """Разбор текста без живой модели: дыры не дописываются."""
    return SupplyContractService(
        db_path=db_path,
        today=date(2026, 10, 20),
        card_recognizer=_block_card_vision,
    )


def test_parse_contract_card_upload_returns_fields_and_writes_no_row(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    client.app.dependency_overrides[get_supply_contract_service] = lambda: _parse_service(db_path)
    manager = _cookie(2, "manager", "manager")
    admin = _cookie(1, "admin", "admin")

    for cookies in (manager, admin):
        response = client.post(
            f"{_CARDS}/parse",
            files={"file": ("card.xlsx", _card_xlsx(), _XLSX)},
            cookies=cookies,
        )
        assert response.status_code == 200
        body = response.json()
        assert body["fields"]["inn"] == "7604010011"
        assert body["fields"]["account"] == "40702810000000000007"

    ensure_schema(db_path)
    assert _count(db_path, "supply_contract") == 0
    assert _count(db_path, "counterparties") == 0


def test_parse_contract_card_upload_rejects_a_non_card_and_writes_no_row(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    client.app.dependency_overrides[get_supply_contract_service] = lambda: _parse_service(db_path)
    manager = _cookie(2, "manager", "manager")

    response = client.post(
        f"{_CARDS}/parse",
        files={"file": ("passport.xlsx", _passport_xlsx(), _XLSX)},
        cookies=manager,
    )

    assert response.status_code == 400
    assert response.json()["detail"] == "это не карточка контрагента"
    ensure_schema(db_path)
    assert _count(db_path, "supply_contract") == 0


def test_parse_contract_card_upload_is_closed_without_a_writer_role(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    shop = _cookie(3, "shop", "production")

    anonymous = client.post(
        f"{_CARDS}/parse",
        files={"file": ("card.xlsx", _card_xlsx(), _XLSX)},
    )
    forbidden = client.post(
        f"{_CARDS}/parse",
        files={"file": ("card.xlsx", _card_xlsx(), _XLSX)},
        cookies=shop,
    )

    assert anonymous.status_code == 401
    assert forbidden.status_code == 403
    ensure_schema(db_path)
    assert _count(db_path, "supply_contract") == 0
