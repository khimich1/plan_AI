"""Справочник клиентов с пометкой договора и одна запись номера этому клиенту."""

from __future__ import annotations

import inspect
import sqlite3
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
from core.kp_db_schema import ensure_schema
from app.repositories.counterparties_repository import CounterpartiesRepository
from app.repositories.supply_contract_repository import SupplyContractRepository
from app.security.session import create_session_token
from app.services.supply_contract_service import SupplyContractService
from core.supply_contract import CANCELLED_STATUS, INVOICE_DOCUMENT, contract_seq
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient

_DIRECTORY = "/api/v1/commercial/archive/directory"
_CARDS = "/api/v1/commercial/archive/contract-cards"
_TODAY = date(2026, 11, 2)

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


_ADMIN = _cookie(1, "admin", "admin")
_MANAGER = _cookie(2, "manager", "manager")
_SHOP = _cookie(3, "shop", "production")


@pytest.fixture()
def directory(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[CsrfAwareTestClient, str]:
    db_path = str(tmp_path / "directory.db")
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    get_settings.cache_clear()
    assert get_settings().plita_db_path.resolve() == Path(db_path).resolve()
    ensure_schema(db_path)
    patch_auth_users(monkeypatch, _USERS)
    app = create_app()
    app.dependency_overrides[get_supply_contract_service] = lambda: SupplyContractService(
        db_path=db_path,
        today=_TODAY,
    )
    return CsrfAwareTestClient(app), db_path


def _client(
    db_path: str,
    *,
    name: str,
    code: str,
    inn: str | None = None,
    kpp: str | None = None,
    is_client: bool = True,
    is_active: bool = True,
) -> int:
    row = CounterpartiesRepository(db_path=db_path).insert(
        code_1c=code,
        name=name,
        inn=inn,
        kpp=kpp,
        is_client=is_client,
        is_active=is_active,
    )
    return int(row["id"])


def _free_card(
    db_path: str,
    *,
    name: str,
    number: str,
    inn: str | None,
    contract_date: date | None = None,
) -> int:
    row = SupplyContractRepository(db_path=db_path).insert_card(
        imported_name=name,
        number=number,
        user_id=1,
        inn=inn,
        contract_date=contract_date,
    )
    assert row["counterparty_id"] is None
    return int(row["id"])


def _contract_rows(db_path: str) -> list[sqlite3.Row]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return list(conn.execute("SELECT * FROM supply_contract ORDER BY id"))
    finally:
        conn.close()


def _party_count(db_path: str) -> int:
    with sqlite3.connect(db_path) as conn:
        return int(conn.execute("SELECT COUNT(*) FROM counterparties").fetchone()[0])


def _max_seq(db_path: str) -> int:
    highest = 0
    for row in _contract_rows(db_path):
        seq = contract_seq(str(row["number"]))
        if seq is not None and seq > highest:
            highest = seq
    return highest


def _search(client: CsrfAwareTestClient, q: str, cookies: dict[str, str] | None = None):
    return client.get(_DIRECTORY, params={"q": q}, cookies=cookies or _ADMIN)


def test_blank_and_one_character_return_no_rows(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")

    omitted = client.get(_DIRECTORY, cookies=_ADMIN)
    assert omitted.status_code == 200
    assert omitted.json() == {"items": [], "count": 0}

    for query in ("", " ", "У", "у"):
        response = _search(client, query)
        assert response.status_code == 200
        assert response.json() == {"items": [], "count": 0}


def test_name_and_inn_find_a_client_and_the_1c_code_does_not(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(
        db_path,
        name="УНИСТРОЙ",
        code="ZZ-NOT-A-NAME",
        inn="7701000001",
        kpp="770101001",
    )
    _client(db_path, name="Ромашка", code="УНИСТРОЙ", inn="5001000099", kpp="500101001")

    by_name = _search(client, "  уни  ")
    assert by_name.status_code == 200
    body = by_name.json()
    assert body["count"] == 1
    assert body["items"] == [
        {
            "id": client_id,
            "name": "УНИСТРОЙ",
            "inn": "7701000001",
            "kpp": "770101001",
            "code": "ZZ-NOT-A-NAME",
            "contract": None,
            "linkable": None,
            "candidates": [],
        }
    ]

    by_inn = _search(client, "7701000001")
    assert by_inn.status_code == 200
    assert [item["id"] for item in by_inn.json()["items"]] == [client_id]

    by_code = _search(client, "ZZ-NOT-A-NAME")
    assert by_code.status_code == 200
    assert by_code.json() == {"items": [], "count": 0}


def test_search_stops_at_ten_active_clients(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    for index in range(1, 12):
        _client(db_path, name=f"Альфа {index:02d}", code=f"00-0000{index:04d}", inn=f"770100{index:04d}")

    response = _search(client, "Альфа")

    assert response.status_code == 200
    assert response.json()["count"] == 10
    assert len(response.json()["items"]) == 10


def test_manager_does_not_see_a_non_client_and_admin_does(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    buyer_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    supplier_id = _client(
        db_path,
        name="ЩебеньПоставка",
        code="00-00000088",
        inn="5002000002",
        is_client=False,
    )
    _client(db_path, name="Скрытый", code="00-00000099", inn="5003000003", is_active=False)

    manager = _search(client, "щеб", cookies=_MANAGER)
    assert manager.status_code == 200
    assert manager.json()["items"] == []

    admin = _search(client, "щеб")
    assert admin.status_code == 200
    assert [item["id"] for item in admin.json()["items"]] == [supplier_id]

    both = _search(client, "уни", cookies=_MANAGER)
    assert [item["id"] for item in both.json()["items"]] == [buyer_id]
    assert _search(client, "скрыт").json()["items"] == []


def test_row_shows_the_active_number_and_no_linkable_card(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001", kpp="770101001")
    card_id = _free_card(
        db_path,
        name="Старое имя",
        number="0882/09/25",
        inn="7701000001",
        contract_date=date(2025, 9, 8),
    )
    spare_id = _free_card(db_path, name="Ещё одна", number="0100/01/24", inn="7701000001")
    linked = SupplyContractRepository(db_path=db_path).link_counterparty(card_id, client_id, user_id=1)
    assert linked is not None

    response = _search(client, "УНИСТРОЙ")

    assert response.status_code == 200
    item = response.json()["items"][0]
    assert item["contract"] == {"id": card_id, "number": "0882/09/25"}
    assert item["linkable"] is None
    assert item["candidates"] == []
    spare = next(row for row in _contract_rows(db_path) if int(row["id"]) == spare_id)
    assert spare["counterparty_id"] is None


def test_one_free_card_and_one_code_are_linkable(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    card_id = _free_card(db_path, name="Бланк", number="0882/09/25", inn="7701000001")

    item = _search(client, "УНИСТРОЙ").json()["items"][0]

    assert item["id"] == client_id
    assert item["contract"] is None
    assert item["linkable"] == {"id": card_id, "number": "0882/09/25"}
    assert item["candidates"] == [{"id": card_id, "number": "0882/09/25"}]


def test_two_free_cards_and_two_codes_are_not_linkable(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    first_card = _free_card(db_path, name="Первая", number="0100/01/24", inn="7701000001")
    second_card = _free_card(db_path, name="Вторая", number="0882/09/25", inn="7701000001")
    _free_card(db_path, name="Чужой ИНН", number="0009/01/24", inn="5001000099")

    item = _search(client, "УНИСТРОЙ").json()["items"][0]
    assert item["contract"] is None
    assert item["linkable"] is None
    assert item["candidates"] == [
        {"id": first_card, "number": "0100/01/24"},
        {"id": second_card, "number": "0882/09/25"},
    ]

    first = _client(db_path, name="КАРАВАН", code="00-00003032", inn="3528341614")
    second = _client(db_path, name="КАРАВАН ДУБЛЬ", code="00-00003033", inn="3528341614")
    free_id = _free_card(db_path, name="Свободная", number="0300/01/25", inn="3528341614")
    taken = _free_card(db_path, name="Уже занята", number="0001/02/24", inn="3528341614")
    owner = _client(db_path, name="Другой владелец", code="00-00000001", inn="5001000001")
    SupplyContractRepository(db_path=db_path).link_counterparty(taken, owner, user_id=1)

    rows = {item["id"]: item for item in _search(client, "КАРАВАН").json()["items"]}
    expected = [{"id": free_id, "number": "0300/01/25"}]
    assert rows[first]["linkable"] is None
    assert rows[second]["linkable"] is None
    assert rows[first]["contract"] is None
    assert rows[second]["contract"] is None
    assert rows[first]["candidates"] == expected
    assert rows[second]["candidates"] == expected


def test_shop_cannot_search_or_write_the_directory(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")

    assert _search(client, "УНИ", cookies=_SHOP).status_code == 403
    denied = client.post(f"{_DIRECTORY}/{client_id}/contract", cookies=_SHOP)
    denied_blank = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json={"legal_form": "ooo", "full_name": "ООО"},
        cookies=_SHOP,
    )
    assert denied.status_code == 403
    assert denied_blank.status_code == 403
    assert _contract_rows(db_path) == []


def test_empty_post_does_not_allocate_or_insert(
    directory: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, db_path = directory
    client_id = _client(
        db_path,
        name="УНИСТРОЙ",
        code="00-00000077",
        inn="7701000001",
        kpp="770101001",
    )
    foreign_id = _free_card(db_path, name="Чужой", number="1031/10/26", inn="5001000099")
    before_parties = _party_count(db_path)
    before_seq = _max_seq(db_path)

    def _refuse_allocate(self: object, conn: object, contract_date: date) -> str:
        raise AssertionError("allocate")

    monkeypatch.setattr(SupplyContractRepository, "_allocate_number", _refuse_allocate)

    response = client.post(f"{_DIRECTORY}/{client_id}/contract", cookies=_ADMIN)
    both = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "0882/09/25", "contract_id": foreign_id},
        cookies=_ADMIN,
    )

    assert response.status_code == 400
    assert response.json()["detail"] == "Укажите номер или карточку"
    assert both.status_code == 400
    assert both.json()["detail"] == "Укажите номер или карточку"
    rows = _contract_rows(db_path)
    assert len(rows) == 1
    assert int(rows[0]["id"]) == foreign_id
    assert rows[0]["counterparty_id"] is None
    assert _max_seq(db_path) == before_seq
    assert _party_count(db_path) == before_parties


def test_typed_missing_number_is_stored_literally_and_seq_stays(
    directory: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, db_path = directory
    client_id = _client(
        db_path,
        name="УНИСТРОЙ",
        code="00-00000077",
        inn="7701000001",
        kpp="770101001",
    )
    _free_card(db_path, name="Чужой", number="1031/10/26", inn="5001000099")
    before_parties = _party_count(db_path)

    def _refuse_allocate(self: object, conn: object, contract_date: date) -> str:
        raise AssertionError("allocate")

    monkeypatch.setattr(SupplyContractRepository, "_allocate_number", _refuse_allocate)

    first = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": " 0882 / 09 / 25 "},
        cookies=_ADMIN,
    )
    second = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "0882/09/25"},
        cookies=_MANAGER,
    )

    assert first.status_code == 200
    assert first.json()["number"] == "0882/09/25"
    assert second.status_code == 200
    assert second.json() == first.json()
    rows = _contract_rows(db_path)
    assert len(rows) == 2
    created = next(row for row in rows if row["number"] == "0882/09/25")
    assert int(created["id"]) == int(first.json()["id"])
    assert int(created["counterparty_id"]) == client_id
    assert created["imported_name"] == "УНИСТРОЙ"
    assert created["inn"] == "7701000001"
    assert created["status"] == "нет"
    assert created["contract_date"] == "2025-09-01"
    assert created["full_name"] is None
    assert created["account"] is None
    assert created["bank_name"] is None
    assert _max_seq(db_path) == 1031
    assert _party_count(db_path) == before_parties


def test_typed_number_without_a_date_keeps_the_date_empty(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")

    response = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "А-15"},
        cookies=_ADMIN,
    )

    assert response.status_code == 200
    assert response.json()["number"] == "А-15"
    created = _contract_rows(db_path)[0]
    assert created["contract_date"] is None
    assert created["status"] == "нет"
    assert int(created["counterparty_id"]) == client_id


def test_known_free_card_is_linked_by_id_or_number_and_keeps_its_number(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    card_id = _free_card(db_path, name="Старое имя", number="0882/09/25", inn="7701000001")
    before_parties = _party_count(db_path)
    before_seq = _max_seq(db_path)

    by_id = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"contract_id": card_id},
        cookies=_ADMIN,
    )

    assert by_id.status_code == 200
    assert by_id.json() == {"id": card_id, "number": "0882/09/25"}
    rows = _contract_rows(db_path)
    assert len(rows) == 1
    assert int(rows[0]["counterparty_id"]) == client_id
    assert rows[0]["imported_name"] == "Старое имя"
    assert rows[0]["number"] == "0882/09/25"
    assert _max_seq(db_path) == before_seq
    assert _party_count(db_path) == before_parties

    again = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "0882/09/25"},
        cookies=_MANAGER,
    )
    assert again.status_code == 200
    assert again.json() == {"id": card_id, "number": "0882/09/25"}
    assert len(_contract_rows(db_path)) == 1


def test_two_free_cards_are_not_auto_picked_and_the_chosen_id_links_one(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    first = _free_card(db_path, name="Первая", number="0100/01/24", inn="7701000001")
    second = _free_card(db_path, name="Вторая", number="0882/09/25", inn="7701000001")

    empty = client.post(f"{_DIRECTORY}/{client_id}/contract", cookies=_ADMIN)
    chosen = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"contract_id": second},
        cookies=_ADMIN,
    )

    assert empty.status_code == 400
    assert chosen.status_code == 200
    assert chosen.json() == {"id": second, "number": "0882/09/25"}
    rows = {int(row["id"]): row for row in _contract_rows(db_path)}
    assert len(rows) == 2
    assert rows[first]["counterparty_id"] is None
    assert int(rows[second]["counterparty_id"]) == client_id
    assert rows[second]["number"] == "0882/09/25"


def test_two_codes_do_not_receive_the_free_card_without_a_choice(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    chosen = _client(db_path, name="КАРАВАН", code="00-00003032", inn="3528341614")
    other = _client(db_path, name="КАРАВАН ДУБЛЬ", code="00-00003033", inn="3528341614")
    free_id = _free_card(db_path, name="Свободная", number="0882/09/25", inn="3528341614")

    response = client.post(f"{_DIRECTORY}/{chosen}/contract", cookies=_ADMIN)

    assert response.status_code == 400
    rows = {int(row["id"]): row for row in _contract_rows(db_path)}
    assert len(rows) == 1
    assert rows[free_id]["counterparty_id"] is None
    owners = [row["counterparty_id"] for row in rows.values()]
    assert chosen not in {None if owner is None else int(owner) for owner in owners}
    assert other not in {None if owner is None else int(owner) for owner in owners}


def test_contract_id_with_a_different_inn_is_refused(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    card_id = _free_card(db_path, name="Чужой ИНН", number="0882/09/25", inn="5001000099")

    by_id = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"contract_id": card_id},
        cookies=_ADMIN,
    )
    by_number = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "0882/09/25"},
        cookies=_ADMIN,
    )
    missing = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"contract_id": 999999},
        cookies=_ADMIN,
    )

    assert by_id.status_code == 400
    assert by_id.json()["detail"] == "ИНН карточки не совпадает с клиентом"
    assert by_number.status_code == 400
    assert by_number.json()["detail"] == "ИНН карточки не совпадает с клиентом"
    assert missing.status_code == 404
    row = _contract_rows(db_path)[0]
    assert row["counterparty_id"] is None
    assert row["number"] == "0882/09/25"


def test_existing_active_contract_is_returned_without_a_second_row(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    card_id = _free_card(db_path, name="Уже есть", number="0882/09/25", inn="7701000001")
    SupplyContractRepository(db_path=db_path).link_counterparty(card_id, client_id, user_id=1)

    response = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "0882/09/25"},
        cookies=_ADMIN,
    )

    assert response.status_code == 200
    assert response.json() == {"id": card_id, "number": "0882/09/25"}
    other = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "0100/01/24"},
        cookies=_ADMIN,
    )
    assert other.status_code == 400
    assert other.json()["detail"] == "У контрагента уже есть договор 0882/09/25"
    assert len(_contract_rows(db_path)) == 1


def test_inactive_and_missing_ids_write_nothing(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    inactive_id = _client(
        db_path,
        name="Бывший",
        code="00-00000099",
        inn="5003000003",
        is_active=False,
    )
    hidden_id = _client(
        db_path,
        name="ЩебеньПоставка",
        code="00-00000088",
        inn="5002000002",
        is_client=False,
    )

    inactive = client.post(f"{_DIRECTORY}/{inactive_id}/contract", cookies=_ADMIN)
    missing = client.post(f"{_DIRECTORY}/999999/contract", cookies=_ADMIN)
    hidden = client.post(f"{_DIRECTORY}/{hidden_id}/contract", cookies=_MANAGER)

    assert inactive.status_code == 400
    assert missing.status_code == 404
    assert hidden.status_code == 404
    assert _contract_rows(db_path) == []


def test_cancelled_number_is_not_replaced_by_an_empty_post(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    card_id = _free_card(db_path, name="Старый", number="0882/09/25", inn="7701000001")
    SupplyContractRepository(db_path=db_path).link_counterparty(card_id, client_id, user_id=1)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE supply_contract SET status = ? WHERE id = ?",
            (CANCELLED_STATUS, card_id),
        )

    empty = client.post(f"{_DIRECTORY}/{client_id}/contract", cookies=_ADMIN)
    repeat = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"number": "0882/09/25"},
        cookies=_ADMIN,
    )

    assert empty.status_code == 400
    assert repeat.status_code == 200
    assert repeat.json() == {"id": card_id, "number": "0882/09/25"}
    rows = _contract_rows(db_path)
    assert len(rows) == 1
    assert rows[0]["status"] == CANCELLED_STATUS
    assert int(rows[0]["counterparty_id"]) == client_id
    assert _max_seq(db_path) == 882


def test_another_clients_number_is_refused_without_a_write(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    owner = _client(db_path, name="Уже связан", code="00-00000001", inn="5001000001")
    buyer = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    owned = _free_card(db_path, name="Чужой договор", number="0001/10/26", inn="5001000001")
    SupplyContractRepository(db_path=db_path).link_counterparty(owned, owner, user_id=1)
    before_seq = _max_seq(db_path)

    by_number = client.post(
        f"{_DIRECTORY}/{buyer}/contract",
        json={"number": "0001/10/26"},
        cookies=_ADMIN,
    )
    by_id = client.post(
        f"{_DIRECTORY}/{buyer}/contract",
        json={"contract_id": owned},
        cookies=_ADMIN,
    )

    assert by_number.status_code == 400
    assert by_number.json()["detail"] == "Номер уже у другого клиента"
    assert by_id.status_code == 400
    assert by_id.json()["detail"] == "Номер уже у другого клиента"
    rows = _contract_rows(db_path)
    assert len(rows) == 1
    assert int(rows[0]["counterparty_id"]) == owner
    assert rows[0]["number"] == "0001/10/26"
    assert _max_seq(db_path) == before_seq


def test_empty_linked_card_clears_the_missing_contract_block_for_that_kp(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import MSG_NEED_CONTRACT, ArchiveService
    from core.supply_contract import attach_contract_number

    client, db_path = directory
    client_id = _client(db_path, name="КАРАВАН ООО", code="00-00003032", inn="3528341614")
    card_id = _free_card(
        db_path,
        name="Старое имя",
        number="0882/09/25",
        inn="3528341614",
        contract_date=date(2025, 9, 8),
    )
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO KP_offers (
                kp_id, creation_date, customer_name, manager_name,
                discount_percent, subtotal, vat_amount, total_amount,
                logistics_cost, counterparty_id, customer_inn, customer_kpp
            ) VALUES (
                15, '11.09.2026', 'КАРАВАН на карточке', 'Иван Иванов',
                0, 10, 0, 10,
                0, ?, '3528341614', '352801001'
            )
            """,
            (client_id,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (15, 'в архиве', 1)"
        )
    export_dir = Path(db_path).parent / "exchange"
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=export_dir.parent / "out",
        export_dir=export_dir,
    )
    user = {"id": 1, "role": "admin"}
    assert service.get_details(15, user=user).invoice_export_block == MSG_NEED_CONTRACT

    empty = client.post(f"{_DIRECTORY}/{client_id}/contract", cookies=_ADMIN)
    assert empty.status_code == 400
    assert service.get_details(15, user=user).invoice_export_block == MSG_NEED_CONTRACT

    response = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"contract_id": card_id},
        cookies=_ADMIN,
    )

    assert response.status_code == 200
    assert response.json() == {"id": card_id, "number": "0882/09/25"}
    created = next(row for row in _contract_rows(db_path) if int(row["id"]) == card_id)
    assert created["full_name"] is None
    assert created["account"] is None
    assert created["bank_name"] is None
    assert service.get_details(15, user=user).invoice_export_block is None
    stamped = attach_contract_number(
        {"Документ": INVOICE_DOCUMENT},
        client_id,
        SupplyContractRepository(db_path=db_path).get_active_by_counterparty(client_id),
    )
    assert stamped["НомерДоговора"] == "0882/09/25"
    assert list(export_dir.glob("*.json")) == []


def _blank_body(**overrides: object) -> dict[str, object]:
    data: dict[str, object] = {
        "legal_form": "ooo",
        "full_name": "ООО УНИСТРОЙ",
        "short_name": "УНИСТРОЙ",
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


def _invoice_service(db_path: str):
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveService

    export_dir = Path(db_path).parent / "exchange"
    return ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=export_dir.parent / "out",
        export_dir=export_dir,
    ), export_dir


def _seed_archived_kp(db_path: str, client_id: int) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO KP_offers (
                kp_id, creation_date, customer_name, manager_name,
                discount_percent, subtotal, vat_amount, total_amount,
                logistics_cost, counterparty_id, customer_inn, customer_kpp
            ) VALUES (
                15, '11.09.2026', 'КАРАВАН на карточке', 'Иван Иванов',
                0, 10, 0, 10,
                0, ?, '3528341614', '352801001'
            )
            """,
            (client_id,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (15, 'в архиве', 1)"
        )


def test_incomplete_blank_does_not_insert_or_move_seq(
    directory: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    _free_card(db_path, name="Чужой", number="1031/10/26", inn="5001000099")

    def _refuse_allocate(self: object, conn: object, contract_date: date) -> str:
        raise AssertionError("allocate")

    monkeypatch.setattr(SupplyContractRepository, "_allocate_number", _refuse_allocate)
    missing_kpp = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json=_blank_body(kpp=None),
        cookies=_ADMIN,
    )
    missing_name = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json={key: value for key, value in _blank_body().items() if key != "full_name"},
        cookies=_ADMIN,
    )

    assert missing_kpp.status_code == 422
    assert missing_kpp.json()["detail"] == "Для ООО и АО укажите КПП"
    assert missing_name.status_code == 422
    rows = _contract_rows(db_path)
    assert len(rows) == 1
    assert rows[0]["number"] == "1031/10/26"
    assert _max_seq(db_path) == 1031


def test_first_ready_blank_issues_one_number_and_repeat_keeps_it(
    directory: tuple[CsrfAwareTestClient, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.archive_service import MSG_NEED_CONTRACT
    from core.supply_contract import MissingSupplyContractError, attach_contract_number

    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    _free_card(db_path, name="Чужой", number="1031/10/26", inn="5001000099")
    _seed_archived_kp(db_path, client_id)
    service, export_dir = _invoice_service(db_path)
    user = {"id": 1, "role": "admin"}
    assert service.get_details(15, user=user).invoice_export_block == MSG_NEED_CONTRACT
    with pytest.raises(MissingSupplyContractError, match="Нет действующего договора поставки"):
        attach_contract_number(
            {"Документ": INVOICE_DOCUMENT},
            client_id,
            SupplyContractRepository(db_path=db_path).get_active_by_counterparty(client_id),
        )

    first = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json=_blank_body(),
        cookies=_ADMIN,
    )

    assert first.status_code == 200
    assert first.json()["number"] == "1032/11/26"
    assert first.json()["requisites_complete"] is True
    created_id = int(first.json()["id"])
    rows = _contract_rows(db_path)
    assert len(rows) == 2
    created = next(row for row in rows if int(row["id"]) == created_id)
    assert int(created["counterparty_id"]) == client_id
    assert created["imported_name"] == "УНИСТРОЙ"
    assert created["status"] == "нет"
    assert created["contract_date"] == "2026-11-02"
    assert created["full_name"] == "ООО УНИСТРОЙ"
    assert created["kpp"] == "770101001"
    assert service.get_details(15, user=user).invoice_export_block is None
    stamped = attach_contract_number(
        {"Документ": INVOICE_DOCUMENT},
        client_id,
        SupplyContractRepository(db_path=db_path).get_active_by_counterparty(client_id),
    )
    assert stamped["НомерДоговора"] == "1032/11/26"
    assert list(export_dir.glob("*.json")) == []

    def _refuse_allocate(self: object, conn: object, contract_date: date) -> str:
        raise AssertionError("allocate")

    monkeypatch.setattr(SupplyContractRepository, "_allocate_number", _refuse_allocate)
    second = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json=_blank_body(short_name="УНИ"),
        cookies=_MANAGER,
    )

    assert second.status_code == 200
    assert second.json()["id"] == created_id
    assert second.json()["number"] == "1032/11/26"
    stored = {int(row["id"]): row for row in _contract_rows(db_path)}
    assert len(stored) == 2
    assert stored[created_id]["number"] == "1032/11/26"
    assert stored[created_id]["short_name"] == "УНИ"
    assert stored[created_id]["contract_date"] == "2026-11-02"


def test_cancelled_number_stays_occupied_for_the_next_blank(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    _free_card(db_path, name="Чужой", number="1031/10/26", inn="5001000099")
    first = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json=_blank_body(),
        cookies=_ADMIN,
    )
    assert first.status_code == 200
    created_id = int(first.json()["id"])
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE supply_contract SET status = ? WHERE id = ?",
            (CANCELLED_STATUS, created_id),
        )

    second = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json=_blank_body(),
        cookies=_ADMIN,
    )

    assert second.status_code == 200
    assert second.json()["number"] == "1033/11/26"
    assert int(second.json()["id"]) != created_id
    rows = {row["number"]: row for row in _contract_rows(db_path)}
    assert rows["1032/11/26"]["status"] == CANCELLED_STATUS
    assert int(rows["1032/11/26"]["counterparty_id"]) == client_id
    assert int(rows["1033/11/26"]["counterparty_id"]) == client_id
    assert rows["1031/10/26"]["counterparty_id"] is None


def test_ready_blank_keeps_the_number_of_an_already_linked_card(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    client_id = _client(db_path, name="УНИСТРОЙ", code="00-00000077", inn="7701000001")
    card_id = _free_card(
        db_path,
        name="Старое имя",
        number="0882/09/25",
        inn="7701000001",
        contract_date=date(2025, 9, 8),
    )
    linked = client.post(
        f"{_DIRECTORY}/{client_id}/contract",
        json={"contract_id": card_id},
        cookies=_ADMIN,
    )
    assert linked.status_code == 200

    saved = client.post(
        f"{_DIRECTORY}/{client_id}/contract-blank",
        json=_blank_body(),
        cookies=_ADMIN,
    )

    assert saved.status_code == 200
    assert saved.json()["id"] == card_id
    assert saved.json()["number"] == "0882/09/25"
    rows = _contract_rows(db_path)
    assert len(rows) == 1
    assert rows[0]["number"] == "0882/09/25"
    assert rows[0]["full_name"] == "ООО УНИСТРОЙ"
    assert rows[0]["contract_date"] == "2025-09-08"
    assert _max_seq(db_path) == 882


def test_contract_card_import_path_still_writes_an_empty_counterparty_id(
    directory: tuple[CsrfAwareTestClient, str],
) -> None:
    client, db_path = directory
    before_parties = _party_count(db_path)

    created = client.post(
        _CARDS,
        json={"name": "Секлюцкий", "inn": "440117229007"},
        cookies=_ADMIN,
    )
    listed = client.get(_CARDS, params={"q": "Секлюц"}, cookies=_ADMIN)

    assert created.status_code == 200
    assert created.json()["counterparty_id"] is None
    assert created.json()["number"] == "0001/11/26"
    assert listed.status_code == 200
    assert listed.json()["count"] == 1
    assert listed.json()["items"][0]["counterparty_id"] is None
    stored = _contract_rows(db_path)
    assert len(stored) == 1
    assert stored[0]["counterparty_id"] is None
    assert _party_count(db_path) == before_parties
