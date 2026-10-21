"""Поиск клиента в мастере КП: карточки договоров, не весь справочник 1С."""

from __future__ import annotations

import sqlite3
from typing import Any

import pytest

from app.repositories.counterparties_repository import CounterpartiesRepository
from app.repositories.supply_contract_repository import SupplyContractRepository
from tests.helpers import kp_db_fixtures as fx
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

WIZARD = "/api/v1/counterparties/contract-cards"
DIRECTORY = "/api/v1/counterparties/search"

TEST_USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 2,
        "username": "prod_user",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


def _count_counterparties(db_path: str) -> int:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute("SELECT COUNT(*) FROM counterparties").fetchone()
    assert row is not None
    return int(row[0])


def _manager() -> dict[str, str]:
    return session_cookie(3, "manager", "manager_a")


def _production() -> dict[str, str]:
    return session_cookie(2, "production", "prod_user")


def _wizard(client: Any, q: str):
    return client.get(WIZARD, params={"q": q}, cookies=_manager())


@pytest.fixture()
def wizard_cards(tmp_path, monkeypatch: pytest.MonkeyPatch):
    from app.core.settings import get_settings
    from app.main import create_app
    from tests.helpers.auth_fixtures import patch_auth_users
    from tests.helpers.csrf import CsrfAwareTestClient

    db_path = fx.make_iso_db(tmp_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db_path)
    monkeypatch.setenv("PB_DB_PATH", db_path)
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)

    parties = CounterpartiesRepository(db_path=db_path)
    directory_only = parties.insert(
        code_1c="00-only",
        name="Только Справочник ООО",
        inn="7703000003",
        kpp="770301001",
        is_client=True,
        source="import",
    )
    linked_party = parties.insert(
        code_1c="00-roma",
        name="РОМАШКА ООО",
        inn="7701000001",
        kpp="770101001",
        is_client=True,
        source="import",
    )
    inactive = parties.insert(
        code_1c="00-off",
        name="АРХИВ ООО",
        inn="8800000000",
        kpp="880001001",
        is_client=True,
        is_active=False,
        source="import",
    )
    supplier = parties.insert(
        code_1c="00-sup",
        name="ПОСТАВЩИК ООО",
        inn="9900000000",
        is_client=False,
        source="import",
    )

    cards = SupplyContractRepository(db_path=db_path)
    unlinked = cards.insert_card(
        imported_name="ИП Секлюцкий",
        number="0001/10/26",
        user_id=1,
        inn="5904111223",
    )
    cards.insert_card(
        imported_name="Другой завод",
        number="0002/11/26",
        user_id=1,
        inn="5904000000",
    )
    linked_card = cards.insert_card(
        imported_name="Карточка Ромашка",
        number="0003/10/26",
        user_id=1,
        inn="1111111111",
    )
    cards.link_counterparty(int(linked_card["id"]), int(linked_party["id"]), user_id=1)
    inactive_card = cards.insert_card(
        imported_name="Спящий клиент",
        number="0004/10/26",
        user_id=1,
        inn="8800000001",
    )
    cards.link_counterparty(int(inactive_card["id"]), int(inactive["id"]), user_id=1)
    supplier_card = cards.insert_card(
        imported_name="Поставщик на карточке",
        number="0005/10/26",
        user_id=1,
        inn="9900000001",
    )
    cards.link_counterparty(int(supplier_card["id"]), int(supplier["id"]), user_id=1)
    nikolsky_party = parties.insert(
        code_1c="00-nik",
        name="НИКОЛЬСКИЙ ЮРИДИЧЕСКОЕ ООО",
        inn="9998887776",
        kpp="999801001",
        is_client=True,
        source="import",
    )
    nikolsky = cards.insert_card(
        imported_name="ООО Никольский",
        number="0007/10/26",
        user_id=1,
        inn="5905123456",
    )
    cards.link_counterparty(int(nikolsky["id"]), int(nikolsky_party["id"]), user_id=1)
    missing_card = cards.insert_card(
        imported_name="Битая связь",
        number="0006/10/26",
        user_id=1,
        inn="1234567890",
    )
    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute(
            "UPDATE supply_contract SET counterparty_id = ? WHERE id = ?",
            (999999, int(missing_card["id"])),
        )
        conn.commit()

    client = CsrfAwareTestClient(create_app())
    seeded = {
        "db": db_path,
        "client": client,
        "directory_only": directory_only,
        "linked_party": linked_party,
        "nikolsky_party": nikolsky_party,
        "nikolsky": nikolsky,
        "unlinked": unlinked,
        "linked_card": linked_card,
        "count": _count_counterparties(db_path),
    }
    yield seeded
    get_settings.cache_clear()


def test_wizard_card_search_finds_unlinked_card_not_directory_only(wizard_cards) -> None:
    client = wizard_cards["client"]
    response = _wizard(client, "секлю")

    assert response.status_code == 200
    body = response.json()
    assert body["count"] == 1
    assert body["items"] == [
        {
            "id": wizard_cards["unlinked"]["id"],
            "name": "ИП Секлюцкий",
            "inn": "5904111223",
            "number": "0001/10/26",
            "requisites_complete": False,
            "counterparty": None,
        }
    ]
    assert "Только Справочник ООО" not in response.text

    missed = _wizard(client, "справочник")
    assert missed.status_code == 200
    assert missed.json() == {"items": [], "count": 0}

    directory = client.get(DIRECTORY, params={"q": "справочник"}, cookies=_manager())
    assert directory.status_code == 200
    assert any(
        item["name"] == "Только Справочник ООО" and item["id"] == wizard_cards["directory_only"]["id"]
        for item in directory.json()["items"]
    )
    assert _count_counterparties(wizard_cards["db"]) == wizard_cards["count"]


def test_wizard_card_search_matches_inn_number_and_ignores_case(wizard_cards) -> None:
    client = wizard_cards["client"]
    by_inn = _wizard(client, "5904111223")
    by_number = _wizard(client, "0001/10")
    by_case = _wizard(client, "СЕКЛЮ")

    assert by_inn.status_code == 200
    assert by_number.status_code == 200
    assert by_case.status_code == 200
    assert [item["number"] for item in by_inn.json()["items"]] == ["0001/10/26"]
    assert [item["number"] for item in by_number.json()["items"]] == ["0001/10/26"]
    assert [item["name"] for item in by_case.json()["items"]] == ["ИП Секлюцкий"]
    assert _count_counterparties(wizard_cards["db"]) == wizard_cards["count"]


def test_wizard_card_search_does_not_treat_percent_or_underscore_as_wildcards(
    wizard_cards,
) -> None:
    client = wizard_cards["client"]
    for query in ("%", "_", "%%", "__", "%секлю", "_еклюц"):
        response = _wizard(client, query)
        assert response.status_code == 200, query
        assert response.json()["items"] == [], query
    assert _count_counterparties(wizard_cards["db"]) == wizard_cards["count"]


def test_wizard_card_search_embeds_active_client_without_insert(wizard_cards) -> None:
    response = _wizard(wizard_cards["client"], "карточка р")

    assert response.status_code == 200
    body = response.json()
    assert body["count"] == 1
    item = body["items"][0]
    assert item["name"] == "Карточка Ромашка"
    assert item["inn"] == "1111111111"
    assert item["number"] == "0003/10/26"
    assert item["requisites_complete"] is False
    assert item["counterparty"] == {
        "id": wizard_cards["linked_party"]["id"],
        "code_1c": "00-roma",
        "name": "РОМАШКА ООО",
        "inn": "7701000001",
        "kpp": "770101001",
    }
    assert _count_counterparties(wizard_cards["db"]) == wizard_cards["count"]


def test_wizard_card_search_omits_inactive_supplier_and_missing_links(wizard_cards) -> None:
    client = wizard_cards["client"]
    for query, name in (
        ("спящий", "Спящий клиент"),
        ("поставщик", "Поставщик на карточке"),
        ("битая", "Битая связь"),
    ):
        response = _wizard(client, query)
        assert response.status_code == 200, query
        body = response.json()
        assert body["count"] == 1, query
        assert body["items"][0]["name"] == name
        assert body["items"][0]["counterparty"] is None
        assert body["items"][0]["requisites_complete"] is False
    assert _count_counterparties(wizard_cards["db"]) == wizard_cards["count"]


def test_wizard_card_search_nikolsky_shows_card_name_not_directory_name(wizard_cards) -> None:
    response = _wizard(wizard_cards["client"], "николь")

    assert response.status_code == 200
    body = response.json()
    assert body["count"] == 1
    item = body["items"][0]
    assert item["name"] == "ООО Никольский"
    assert item["inn"] == "5905123456"
    assert item["number"] == "0007/10/26"
    assert item["counterparty"]["id"] == wizard_cards["nikolsky_party"]["id"]
    assert item["counterparty"]["code_1c"] == "00-nik"
    assert item["counterparty"]["name"] == "НИКОЛЬСКИЙ ЮРИДИЧЕСКОЕ ООО"
    assert item["name"] != item["counterparty"]["name"]
    assert _count_counterparties(wizard_cards["db"]) == wizard_cards["count"]


def test_wizard_card_search_one_character_returns_empty(wizard_cards) -> None:
    response = _wizard(wizard_cards["client"], "с")

    assert response.status_code == 200
    assert response.json() == {"items": [], "count": 0}


def test_wizard_card_search_forbidden_for_production(wizard_cards) -> None:
    response = wizard_cards["client"].get(
        WIZARD, params={"q": "секлю"}, cookies=_production()
    )
    assert response.status_code == 403
