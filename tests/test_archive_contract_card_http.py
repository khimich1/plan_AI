"""HTTP карточки договора: поиск, создание, чтение и правка (CDC-008)."""

from __future__ import annotations

import inspect
import sqlite3
from datetime import date
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
from app.dependencies.services import get_supply_contract_service
from app.main import create_app
from app.security.session import create_session_token
from app.services.supply_contract_service import SupplyContractService
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient

_CARDS = "/api/v1/commercial/archive/contract-cards"
_CONTRACTS = "/api/v1/commercial/archive/supply-contracts"

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


def _cookie(user_id: int, username: str, role: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }


@pytest.fixture()
def cards(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[CsrfAwareTestClient, str]:
    db_path = str(tmp_path / "cards.db")
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    assert get_settings().plita_db_path.resolve() == Path(db_path).resolve()
    patch_auth_users(monkeypatch, _USERS)
    app = create_app()
    scans_dir = tmp_path / "scans"
    app.dependency_overrides[get_supply_contract_service] = lambda: SupplyContractService(
        db_path=db_path,
        today=date(2026, 10, 20),
        scans_dir=scans_dir,
    )
    return CsrfAwareTestClient(app), db_path


def test_blank_search_lists_every_card_and_one_character_does_not(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = client.post(
        _CARDS,
        json={"name": "Секлюцкий", "number": "0793/02/25"},
        cookies=admin,
    )
    assert created.status_code == 200

    for query in ("", " "):
        response = client.get(_CARDS, params={"q": query}, cookies=admin)
        assert response.status_code == 200
        body = response.json()
        assert body["count"] == 1
        assert body["items"][0]["number"] == "0793/02/25"

    omitted = client.get(_CARDS, cookies=admin)
    assert omitted.status_code == 200
    assert omitted.json()["count"] == 1

    for query in ("с", "С"):
        response = client.get(_CARDS, params={"q": query}, cookies=admin)
        assert response.status_code == 200
        assert response.json() == {"items": [], "count": 0}


def test_search_finds_card_by_casefolded_name(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    client.post(
        _CARDS,
        json={"name": "ИП Секлюцкий", "inn": "440117229007", "number": "0793/02/25"},
        cookies=admin,
    )
    client.post(
        _CARDS,
        json={"name": "Трест", "number": "1000/10/26"},
        cookies=admin,
    )

    response = client.get(_CARDS, params={"q": "  СЕКЛЮЦ  "}, cookies=admin)

    assert response.status_code == 200
    body = response.json()
    assert body["count"] == 1
    assert [item["number"] for item in body["items"]] == ["0793/02/25"]
    assert body["items"][0]["name"] == "ИП Секлюцкий"
    assert body["items"][0]["inn"] == "440117229007"


def test_create_without_number_returns_the_service_card(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards

    response = client.post(
        _CARDS,
        json={"name": "  Новый контрагент  "},
        cookies=_cookie(1, "admin", "admin"),
    )

    assert response.status_code == 200
    body = response.json()
    assert body["number"] == "0001/10/26"
    assert body["contract_date"] == "2026-10-20"
    assert body["name"] == "Новый контрагент"
    assert body["imported_name"] == "Новый контрагент"
    assert body["inn"] is None
    assert body["legal_form"] is None
    assert body["full_name"] is None
    assert body["manager_name"] is None
    assert body["email"] is None
    assert body["status"] == "нет"
    assert body["counterparty_id"] is None
    assert "scan_path" not in body


def test_get_card_returns_the_created_row(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = client.post(
        _CARDS,
        json={"name": "ИП Секлюцкий", "inn": "440117229007", "number": "0786/1/25"},
        cookies=admin,
    )
    card_id = created.json()["id"]

    response = client.get(f"{_CARDS}/{card_id}", cookies=admin)

    assert response.status_code == 200
    body = response.json()
    assert body["id"] == card_id
    assert body["number"] == "0786/1/25"
    assert body["contract_date"] == "2025-01-01"
    assert body["name"] == "ИП Секлюцкий"
    assert body["inn"] == "440117229007"
    assert body["legal_form"] is None
    assert body["ogrn"] is None


def _buyer() -> dict[str, object]:
    return {
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


def test_search_marks_blank_card_incomplete_and_saved_requisites_complete(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    empty = client.post(
        _CARDS,
        json={"name": "Импорт", "number": "0793/02/25"},
        cookies=admin,
    )
    filled = client.post(
        _CARDS,
        json={"name": "Ромашка", "number": "0001/10/26"},
        cookies=admin,
    )
    assert empty.status_code == 200
    assert filled.status_code == 200
    saved = client.post(
        f"{_CONTRACTS}/{filled.json()['id']}/requisites",
        json=_buyer(),
        cookies=admin,
    )
    assert saved.status_code == 200

    response = client.get(_CARDS, params={"q": ""}, cookies=admin)
    opened = client.get(f"{_CARDS}/{empty.json()['id']}", cookies=admin)

    assert response.status_code == 200
    assert opened.status_code == 200
    by_number = {item["number"]: item["requisites_complete"] for item in response.json()["items"]}
    assert by_number["0793/02/25"] is False
    assert by_number["0001/10/26"] is True
    assert opened.json()["requisites_complete"] is False
    assert saved.json()["requisites_complete"] is True


def test_get_missing_card_is_404(cards: tuple[CsrfAwareTestClient, str]) -> None:
    client, _db_path = cards

    response = client.get(f"{_CARDS}/404", cookies=_cookie(1, "admin", "admin"))

    assert response.status_code == 404
    assert response.json()["detail"] == "Договор №404 не найден"


def test_create_with_taken_number_is_400(cards: tuple[CsrfAwareTestClient, str]) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    client.post(
        _CARDS,
        json={"name": "ИП Секлюцкий", "number": "0786/1/25"},
        cookies=admin,
    )

    response = client.post(
        _CARDS,
        json={"name": "Другой", "number": "0786 / 1 / 25"},
        cookies=admin,
    )

    assert response.status_code == 400
    assert response.json()["detail"] == "Номер 0786/1/25 уже есть"


def test_patch_to_taken_number_is_400_and_keeps_the_row(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    first = client.post(
        _CARDS,
        json={"name": "ИП Секлюцкий", "number": "0786/1/25"},
        cookies=admin,
    )
    client.post(
        _CARDS,
        json={"name": "Трест", "number": "0793/02/25"},
        cookies=admin,
    )
    card_id = first.json()["id"]

    response = client.patch(
        f"{_CARDS}/{card_id}",
        json={"number": "0793 / 02 / 25", "name": "Не должно сохраниться"},
        cookies=admin,
    )

    assert response.status_code == 400
    assert response.json()["detail"] == "Номер 0793/02/25 уже есть"
    kept = client.get(f"{_CARDS}/{card_id}", cookies=admin)
    assert kept.status_code == 200
    assert kept.json()["number"] == "0786/1/25"
    assert kept.json()["name"] == "ИП Секлюцкий"
    assert kept.json()["contract_date"] == "2025-01-01"


def test_patch_free_number_and_name_keeps_the_date(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = client.post(
        _CARDS,
        json={"name": "ИП Секлюцкий", "inn": "440117229007", "number": "0786/1/25"},
        cookies=admin,
    )
    card_id = created.json()["id"]

    response = client.patch(
        f"{_CARDS}/{card_id}",
        json={"name": "  ИП Секлюцкий А.А.  ", "number": "1000/10/26", "inn": "7702000002"},
        cookies=admin,
    )

    assert response.status_code == 200
    body = response.json()
    assert body["number"] == "1000/10/26"
    assert body["contract_date"] == "2025-01-01"
    assert body["name"] == "ИП Секлюцкий А.А."
    assert body["inn"] == "7702000002"


def test_anonymous_card_routes_are_401(cards: tuple[CsrfAwareTestClient, str]) -> None:
    client, _db_path = cards

    assert client.get(_CARDS, params={"q": "секлюц"}).status_code == 401
    assert client.post(_CARDS, json={"name": "Ромашка"}).status_code == 401
    assert client.get(f"{_CARDS}/1").status_code == 401
    assert client.patch(f"{_CARDS}/1", json={"name": "Ромашка"}).status_code == 401
    assert client.delete(f"{_CARDS}/1").status_code == 401
    assert client.post(f"{_CARDS}/import", files={"file": ("cards.xlsx", b"x", _XLSX)}).status_code == 401


def test_production_role_cannot_read_or_write_cards(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    shop = _cookie(3, "shop", "production")

    assert client.get(_CARDS, params={"q": "секлюц"}, cookies=shop).status_code == 403
    assert client.post(_CARDS, json={"name": "Ромашка"}, cookies=shop).status_code == 403
    assert client.get(f"{_CARDS}/1", cookies=shop).status_code == 403
    assert client.patch(
        f"{_CARDS}/1",
        json={"name": "Ромашка"},
        cookies=shop,
    ).status_code == 403
    assert client.delete(f"{_CARDS}/1", cookies=shop).status_code == 403
    assert client.post(
        f"{_CARDS}/import",
        files={"file": ("cards.xlsx", b"x", _XLSX)},
        cookies=shop,
    ).status_code == 403


def _seed_duplicate_card(client: CsrfAwareTestClient) -> None:
    created = client.post(
        _CARDS,
        json={"name": "ООО Ромашка", "inn": "7701000001", "number": "0001/10/26"},
        cookies=_cookie(1, "admin", "admin"),
    )
    assert created.status_code == 200
    listed = client.get(
        f"{_CARDS}/duplicates",
        params={"name": "ООО Ромашка", "inn": "7701000001"},
        cookies=_cookie(1, "admin", "admin"),
    )
    assert listed.status_code == 200
    assert listed.json()["count"] == 1
    assert listed.json()["items"][0]["name"] == "ООО Ромашка"
    assert listed.json()["items"][0]["inn"] == "7701000001"
    assert listed.json()["items"][0]["number"] == "0001/10/26"


def test_anonymous_duplicate_search_is_401(cards: tuple[CsrfAwareTestClient, str]) -> None:
    client, _db_path = cards
    _seed_duplicate_card(client)

    hidden = client.get(
        f"{_CARDS}/duplicates",
        params={"name": "ООО Ромашка", "inn": "7701000001"},
    )

    assert hidden.status_code == 401
    assert hidden.json() == {"detail": "Not authenticated"}


def test_production_role_cannot_read_duplicate_cards(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    _seed_duplicate_card(client)
    shop = _cookie(3, "shop", "production")

    hidden = client.get(
        f"{_CARDS}/duplicates",
        params={"name": "ООО Ромашка", "inn": "7701000001"},
        cookies=shop,
    )

    assert hidden.status_code == 403
    assert hidden.json() == {"detail": "Forbidden"}


def test_manager_can_create_and_search(cards: tuple[CsrfAwareTestClient, str]) -> None:
    client, _db_path = cards
    manager = _cookie(2, "manager", "manager")

    created = client.post(_CARDS, json={"name": "Ромашка"}, cookies=manager)
    assert created.status_code == 200
    assert created.json()["number"] == "0001/10/26"

    found = client.get(_CARDS, params={"q": "рома"}, cookies=manager)
    assert found.status_code == 200
    assert [item["number"] for item in found.json()["items"]] == ["0001/10/26"]


def test_patch_status_of_blank_card_is_200_not_500(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, _db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = client.post(
        _CARDS,
        json={"name": "Без реквизитов", "number": "ТБ-1"},
        cookies=admin,
    )
    assert created.status_code == 200
    card_id = created.json()["id"]
    assert created.json()["contract_date"] is None
    assert created.json()["legal_form"] is None

    response = client.patch(
        f"{_CONTRACTS}/{card_id}",
        json={"status": "подписан по ЭДО"},
        cookies=admin,
    )

    assert response.status_code == 200
    body = response.json()
    assert body["status"] == "подписан по ЭДО"
    assert body["number"] == "ТБ-1"
    assert body["contract_date"] is None
    assert body["manager_name"] is None
    assert body["legal_form"] is None
    assert body["inn"] is None
    assert body["email"] is None
    assert body["bik"] is None
    assert "scan_path" not in body

    registry = client.get(_CONTRACTS, cookies=admin)
    assert registry.status_code == 200
    row = next(item for item in registry.json() if item["id"] == card_id)
    assert row["contract_date"] is None
    assert row["manager_name"] is None
    assert row["number"] == "ТБ-1"
    assert row["status"] == "подписан по ЭДО"


_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_IMPORT = f"{_CARDS}/import"


def _xlsx_bytes(
    db_path: str,
    rows: list[tuple[str, str, str]],
    *,
    filename: str = "contract-cards.xlsx",
    headers: tuple[str, str, str] = ("Номер договора", "Имя контрагента", "ИНН"),
) -> bytes:
    path = Path(db_path).parent / filename
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = "ИНН без сомнений"
    sheet.append(list(headers))
    for number, name, inn in rows:
        sheet.append([number, name, inn])
    book.save(path)
    return path.read_bytes()


def _upload(
    db_path: str,
    rows: list[tuple[str, str, str]],
    *,
    filename: str = "contract-cards.xlsx",
    headers: tuple[str, str, str] = ("Номер договора", "Имя контрагента", "ИНН"),
) -> dict[str, tuple[str, bytes, str]]:
    payload = _xlsx_bytes(db_path, rows, filename=filename, headers=headers)
    return {"file": (filename, payload, _XLSX)}


def test_import_route_is_declared_above_card_id() -> None:
    from app.api.v1.endpoints import archive as archive_routes

    paths = [route.path for route in archive_routes._archive_router.routes]
    import_path = "/commercial/archive/contract-cards/import"
    card_path = "/commercial/archive/contract-cards/{card_id}"
    assert import_path in paths
    assert paths.index(import_path) < paths.index(card_path)


def test_admin_import_returns_created_and_skipped(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    payload = _upload(
        db_path,
        [
            ("0793/02/25", "ИП Секлюцкий", "440117229007"),
            ("0786/1/25", "ИП Ручной", "5001000001"),
        ],
    )

    response = client.post(_IMPORT, files=payload, cookies=_cookie(1, "admin", "admin"))

    assert response.status_code == 200
    assert response.json() == {"created": 2, "skipped": 0}
    found = client.get(_CARDS, params={"q": "ип"}, cookies=_cookie(1, "admin", "admin"))
    assert found.status_code == 200
    body = found.json()
    assert body["count"] == 2
    by_number = {item["number"]: item for item in body["items"]}
    assert by_number["0793/02/25"]["name"] == "ИП Секлюцкий"
    assert by_number["0793/02/25"]["inn"] == "440117229007"
    assert by_number["0786/1/25"]["name"] == "ИП Ручной"
    assert by_number["0786/1/25"]["contract_date"] == "2025-01-01"
    assert by_number["0786/1/25"]["status"] == "нет"


def test_manager_import_is_403_and_adds_no_row(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    payload = _upload(
        db_path,
        [("0786/1/25", "ИП Ручной", "5001000001")],
        filename="manager.xlsx",
    )

    response = client.post(_IMPORT, files=payload, cookies=_cookie(2, "manager", "manager"))

    assert response.status_code == 403
    found = client.get(_CARDS, params={"q": "0786"}, cookies=_cookie(1, "admin", "admin"))
    assert found.status_code == 200
    assert found.json() == {"items": [], "count": 0}


def test_repeat_import_of_the_same_number_does_not_add_a_row(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    admin = _cookie(1, "admin", "admin")
    first = _upload(
        db_path,
        [("0786/1/25", "ИП Ручной", "5001000001")],
        filename="first.xlsx",
    )
    created = client.post(_IMPORT, files=first, cookies=admin)
    assert created.status_code == 200
    assert created.json() == {"created": 1, "skipped": 0}
    card_id = client.get(_CARDS, params={"q": "0786"}, cookies=admin).json()["items"][0]["id"]
    renamed = client.patch(
        f"{_CARDS}/{card_id}",
        json={"name": "ИП Ручной А.А."},
        cookies=admin,
    )
    assert renamed.status_code == 200

    again = client.post(
        _IMPORT,
        files=_upload(
            db_path,
            [("0786/1/25", "Другое имя", "999999999999")],
            filename="again.xlsx",
        ),
        cookies=admin,
    )

    assert again.status_code == 200
    assert again.json() == {"created": 0, "skipped": 1}
    found = client.get(_CARDS, params={"q": "0786"}, cookies=admin)
    assert found.status_code == 200
    body = found.json()
    assert body["count"] == 1
    assert body["items"][0]["id"] == card_id
    assert body["items"][0]["name"] == "ИП Ручной А.А."
    assert body["items"][0]["inn"] == "5001000001"
    assert body["items"][0]["number"] == "0786/1/25"


def test_import_parse_error_is_400_and_adds_no_row(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    admin = _cookie(1, "admin", "admin")
    payload = _upload(
        db_path,
        [("0786/1/25", "ИП Ручной", "5001000001")],
        filename="broken.xlsx",
        headers=("Номер", "Имя", "Код"),
    )

    response = client.post(_IMPORT, files=payload, cookies=admin)

    assert response.status_code == 400
    assert response.json()["detail"] == (
        "Не найдены колонки: номер договора, имя контрагента, ИНН"
    )
    found = client.get(_CARDS, params={"q": "0786"}, cookies=admin)
    assert found.status_code == 200
    assert found.json() == {"items": [], "count": 0}


def _count(db_path: str, table: str) -> int:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    assert row is not None
    return int(row[0])


def _counterparties(db_path: str) -> list[tuple[str, str, str | None]]:
    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            "SELECT code_1c, name, inn FROM counterparties ORDER BY id"
        ).fetchall()
    return [(str(code), str(name), inn) for code, name, inn in rows]


def _event_contracts(db_path: str) -> list[int]:
    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            "SELECT contract_id FROM supply_contract_event ORDER BY id"
        ).fetchall()
    return [int(row[0]) for row in rows]


def _scan_files(db_path: str) -> list[Path]:
    root = Path(db_path).parent / "scans"
    if not root.exists():
        return []
    return [path for path in root.rglob("*") if path.is_file()]


def _seed_counterparty(db_path: str) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized, inn)
            VALUES ('00-0001', 'ООО Ромашка', 'ооо ромашка', '7701000001')
            """
        )


def test_manager_delete_frees_the_number_and_keeps_counterparties(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    """Слово УДАЛИТЬ живёт только в форме. Запрос удаления его не передаёт."""
    client, db_path = cards
    manager = _cookie(2, "manager", "manager")
    created = client.post(
        _CARDS,
        json={"name": "ИП Секлюцкий", "inn": "440117229007", "number": "0786/1/25"},
        cookies=manager,
    )
    assert created.status_code == 200
    card_id = created.json()["id"]
    other = client.post(
        _CARDS,
        json={"name": "Трест", "number": "1000/10/26"},
        cookies=manager,
    )
    assert other.status_code == 200
    other_id = other.json()["id"]
    renamed = client.patch(
        f"{_CARDS}/{other_id}",
        json={"number": "1001/10/26"},
        cookies=manager,
    )
    assert renamed.status_code == 200
    scanned = client.patch(
        f"{_CONTRACTS}/{card_id}",
        data={"scan_note": "получен скан"},
        files={"file": ("scan.pdf", b"%PDF-1.4 scan", "application/pdf")},
        cookies=manager,
    )
    assert scanned.status_code == 200
    assert scanned.json()["has_scan"] is True
    scan_files = _scan_files(db_path)
    assert len(scan_files) == 1
    assert card_id in _event_contracts(db_path)
    assert other_id in _event_contracts(db_path)
    _seed_counterparty(db_path)
    parties = _counterparties(db_path)
    assert len(parties) == 1

    response = client.delete(f"{_CARDS}/{card_id}", cookies=manager)

    assert response.status_code == 204
    assert response.content == b""
    missing = client.get(f"{_CARDS}/{card_id}", cookies=manager)
    assert missing.status_code == 404
    assert missing.json()["detail"] == f"Договор №{card_id} не найден"
    listed = client.get(_CARDS, cookies=manager)
    assert listed.status_code == 200
    numbers = [item["number"] for item in listed.json()["items"]]
    assert "0786/1/25" not in numbers
    assert numbers == ["1001/10/26"]
    assert _event_contracts(db_path) == [other_id]
    assert _scan_files(db_path) == []
    assert _count(db_path, "counterparties") == 1
    assert _counterparties(db_path) == parties
    again = client.post(
        _CARDS,
        json={"name": "Снова Секлюцкий", "number": "0786/1/25"},
        cookies=manager,
    )
    assert again.status_code == 200
    assert again.json()["number"] == "0786/1/25"
    assert again.json()["id"] != card_id


def test_delete_missing_card_is_404(cards: tuple[CsrfAwareTestClient, str]) -> None:
    client, db_path = cards

    response = client.delete(f"{_CARDS}/404", cookies=_cookie(2, "manager", "manager"))

    assert response.status_code == 404
    assert response.json()["detail"] == "Договор №404 не найден"
    assert _count(db_path, "supply_contract") == 0
    assert _count(db_path, "counterparties") == 0
