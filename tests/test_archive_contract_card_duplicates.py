"""Поиск дубля карточки договора: ИНН, отображаемое имя, пустые, не справочник."""

from __future__ import annotations

import inspect
from datetime import date
from pathlib import Path

import fastapi.dependencies.utils as _fastapi_dependencies
import pytest


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
from core import kp_db_schema
from core.kp_db_common import _connect
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient

_CARDS = "/api/v1/commercial/archive/contract-cards"
_CANCELLED = "отмена не будем работать"

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
]


def _service(tmp_path: Path) -> SupplyContractService:
    db_path = str(tmp_path / "cards.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return SupplyContractService(db_path=db_path, today=date(2026, 10, 20))


def _writer() -> dict:
    return {"id": 7, "role": "manager", "username": "manager"}


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


def _numbers(rows: list[dict]) -> list[str]:
    return [str(row["number"]) for row in rows]


def test_inn_duplicate_is_the_whole_digit_string(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    service.create_card(name="Ромашка", inn="5024243403", number="0001/10/26", user=user)
    service.create_card(name="Другая", inn="7701000001", number="0002/10/26", user=user)
    before = _count(service.repository.db_path, "supply_contract")

    found = service.find_duplicate_cards(name="", inn="5024243403", user=user)
    prefix = service.find_duplicate_cards(name="", inn="5024", user=user)

    assert _numbers(found) == ["0001/10/26"]
    assert found[0]["inn"] == "5024243403"
    assert prefix == []
    assert _count(service.repository.db_path, "supply_contract") == before


def test_inn_separators_match_and_leading_zeros_do_not(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    service.create_card(name="Дефис", inn="5024-243-403", number="0004/10/26", user=user)
    service.create_card(name="Нули", inn="05024243403", number="0005/10/26", user=user)

    spaced = service.find_duplicate_cards(name="", inn="5024 243403", user=user)
    zeros = service.find_duplicate_cards(name="", inn="05024243403", user=user)
    plain = service.find_duplicate_cards(name="", inn="5024243403", user=user)
    short = service.find_duplicate_cards(name="", inn="5024", user=user)

    assert _numbers(spaced) == ["0004/10/26"]
    assert _numbers(zeros) == ["0005/10/26"]
    assert _numbers(plain) == ["0004/10/26"]
    assert short == []


def test_blank_inn_whitespace_and_letters_match_nobody(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    service.create_card(name="Без ИНН", number="0002/10/26", user=user)
    service.create_card(name="Буквы", inn="нет", number="0003/10/26", user=user)
    service.create_card(name="Чужой", inn="7701000001", number="0004/10/26", user=user)

    for inn in ("", "   ", "нет", "—"):
        assert service.find_duplicate_cards(name="", inn=inn, user=user) == []


def test_name_duplicate_is_the_display_name_not_a_substring(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    service.create_card(name="ООО  Ромашка", inn="7701000001", number="0001/10/26", user=user)
    service.create_card(name="Строймонтаж", number="0002/10/26", user=user)
    service.create_card(name="ООО Строй", number="0003/10/26", user=user)
    service.create_card(name="Сёмка", number="0007/10/26", user=user)
    service.create_card(name="ОООРомашка", number="0008/10/26", user=user)
    service.create_card(name='ООО «Ромашка»', number="0012/10/26", user=user)
    zavod = service.create_card(name="Завод", number="0006/10/26", user=user)
    fallback = service.create_card(name="Черновик", number="0013/10/26", user=user)
    _set_columns(
        service.repository.db_path,
        int(zavod["id"]),
        full_name="Ромашка Полная",
        short_name="Ромашка",
    )
    _set_columns(
        service.repository.db_path,
        int(fallback["id"]),
        imported_name="   ",
        full_name="ООО Поле",
        short_name="Поле",
    )

    folded = service.find_duplicate_cards(name="ооо  ромашка", inn="", user=user)
    fragment = service.find_duplicate_cards(name="Строй", inn="", user=user)
    yo = service.find_duplicate_cards(name="Семка", inn="", user=user)
    glued = service.find_duplicate_cards(name="ОООРомашка", inn="", user=user)
    hidden = service.find_duplicate_cards(name="Ромашка", inn="", user=user)
    hidden_full = service.find_duplicate_cards(name="Ромашка Полная", inn="", user=user)
    display = service.find_duplicate_cards(name="Завод", inn="", user=user)
    from_full = service.find_duplicate_cards(name="ооо поле", inn="", user=user)

    assert _numbers(folded) == ["0001/10/26"]
    assert folded[0]["name"] == "ООО  Ромашка"
    assert fragment == []
    assert yo == []
    assert _numbers(glued) == ["0008/10/26"]
    assert hidden == []
    assert hidden_full == []
    assert _numbers(display) == ["0006/10/26"]
    assert _numbers(from_full) == ["0013/10/26"]
    assert from_full[0]["name"] == "ООО Поле"


def test_empty_name_matches_nothing(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    blank = service.create_card(name="Будет пустым", number="0010/10/26", user=user)
    service.create_card(name="ООО Ромашка", number="0001/10/26", user=user)
    _set_columns(
        service.repository.db_path,
        int(blank["id"]),
        imported_name=None,
        full_name=None,
        short_name=None,
    )
    before = _count(service.repository.db_path, "supply_contract")

    assert service.find_duplicate_cards(name="", inn="", user=user) == []
    assert service.find_duplicate_cards(name="   ", inn="  ", user=user) == []
    assert _count(service.repository.db_path, "supply_contract") == before


def test_card_matching_both_keys_is_listed_once_in_id_order(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    service.create_card(name="ООО Ромашка", inn="7701000001", number="0001/10/26", user=user)
    service.create_card(name="Другая", inn="7701000001", number="0002/10/26", user=user)
    service.create_card(name="Чужая", inn="5001000000", number="0003/10/26", user=user)

    found = service.find_duplicate_cards(name="ООО Ромашка", inn="7701000001", user=user)

    assert _numbers(found) == ["0001/10/26", "0002/10/26"]
    assert found[0]["id"] < found[1]["id"]


def test_directory_row_is_excluded_and_cancelled_card_is_included(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    service.create_card(name="ООО Ромашка", inn="7701000001", number="0001/10/26", user=user)
    cancelled = service.create_card(name="Отменённая", inn="7702000002", number="0002/10/26", user=user)
    _set_columns(service.repository.db_path, int(cancelled["id"]), status=_CANCELLED)
    with _connect(service.repository.db_path) as conn:
        conn.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized, inn)
            VALUES ('00-0009', 'Только справочник', 'только справочник', '5001000000')
            """
        )
    contracts = _count(service.repository.db_path, "supply_contract")
    parties = _count(service.repository.db_path, "counterparties")

    directory = service.find_duplicate_cards(name="Только справочник", inn="5001000000", user=user)
    same = service.find_duplicate_cards(name="ООО Ромашка", inn="7701000001", user=user)
    dropped = service.find_duplicate_cards(name="Отменённая", inn="", user=user)

    assert directory == []
    assert _numbers(same) == ["0001/10/26"]
    assert _numbers(dropped) == ["0002/10/26"]
    assert dropped[0]["status"] == _CANCELLED
    assert _count(service.repository.db_path, "supply_contract") == contracts
    assert _count(service.repository.db_path, "counterparties") == parties


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
    app.dependency_overrides[get_supply_contract_service] = lambda: SupplyContractService(
        db_path=db_path,
        today=date(2026, 10, 20),
    )
    return CsrfAwareTestClient(app), db_path


def test_get_duplicates_writes_nothing_and_post_still_creates_a_second_card(
    cards: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = cards
    admin = _cookie(1, "admin", "admin")
    created = client.post(
        _CARDS,
        json={"name": "ООО Ромашка", "inn": "7701000001", "number": "0001/10/26"},
        cookies=admin,
    )
    assert created.status_code == 200
    assert _count(db_path, "supply_contract") == 1

    found = client.get(
        f"{_CARDS}/duplicates",
        params={"name": "ооо  ромашка", "inn": "7701 000001"},
        cookies=admin,
    )
    assert found.status_code == 200
    body = found.json()
    assert body["count"] == 1
    assert body["items"][0]["name"] == "ООО Ромашка"
    assert body["items"][0]["inn"] == "7701000001"
    assert body["items"][0]["number"] == "0001/10/26"
    assert body["items"][0]["id"] == created.json()["id"]

    empty = client.get(_CARDS + "/duplicates", params={"name": "", "inn": ""}, cookies=admin)
    assert empty.status_code == 200
    assert empty.json() == {"items": [], "count": 0}
    prefix = client.get(
        f"{_CARDS}/duplicates",
        params={"name": "", "inn": "7701"},
        cookies=admin,
    )
    assert prefix.status_code == 200
    assert prefix.json() == {"items": [], "count": 0}
    assert _count(db_path, "supply_contract") == 1

    second = client.post(
        _CARDS,
        json={"name": "ООО Ромашка", "inn": "7701000001", "number": "0002/10/26"},
        cookies=admin,
    )
    assert second.status_code == 200
    assert second.json()["id"] != created.json()["id"]
    assert second.json()["number"] == "0002/10/26"
    assert _count(db_path, "supply_contract") == 2
