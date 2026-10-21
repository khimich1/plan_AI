"""Карточки контрагентов в архиве: поиск и правка номера договора."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from app.dependencies.services import get_counterparties_service, get_supply_contract_service
from app.main import create_app
from app.repositories.counterparties_repository import CounterpartiesRepository
from app.security.session import create_session_token
from app.services.counterparties_service import CounterpartiesService
from app.services.supply_contract_service import (
    SupplyContractFieldError,
    SupplyContractForbiddenError,
    SupplyContractNotFoundError,
    SupplyContractService,
    SupplyContractValidationError,
)
from core import kp_db_schema
from core.kp_db_common import _connect
from core.supply_contract import CANCELLED_STATUS, normalize_contract_number
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.test_supply_contract import ADMIN, _buyer, _fresh_db, _insert_contract, _seed_counterparty

MANAGER = {"id": 2, "role": "manager", "username": "manager"}


def _cards(db_path: str, q: str) -> list[dict]:
    return CounterpartiesRepository(db_path=db_path).search_cards(q, clients_only=True)


def test_normalize_contract_number_rejects_blank_and_keeps_slashes() -> None:
    assert normalize_contract_number(" 1030/09/26 ") == "1030/09/26"
    with pytest.raises(ValueError):
        normalize_contract_number("без цифр")
    with pytest.raises(ValueError):
        normalize_contract_number("1030 09 26")


def test_search_matches_name_inn_and_contract_number(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "cards.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        romashka = _seed_counterparty(conn, "00-1")
        other = _seed_counterparty(conn, "00-2")
        conn.execute(
            "UPDATE counterparties SET name = ?, name_normalized = ?, inn = ? WHERE id = ?",
            ("ВАСИЛЁК ООО", "василёк ооо", "7802000002", other),
        )
        _insert_contract(conn, counterparty_id=romashka, number="1028/09/26")
        _insert_contract(
            conn,
            counterparty_id=other,
            number="1001/01/26",
            status=CANCELLED_STATUS,
        )

    by_name = _cards(db_path, "ромаш")
    assert [row["id"] for row in by_name] == [romashka]
    assert by_name[0]["contract_number"] == "1028/09/26"

    by_inn = _cards(db_path, "7802000002")
    assert [row["id"] for row in by_inn] == [other]
    assert by_inn[0]["contract_number"] is None

    by_old_number = _cards(db_path, "1001/01")
    assert [row["id"] for row in by_old_number] == [other]

    assert _cards(db_path, "%") == []
    assert _cards(db_path, "%%") == []


def test_search_skips_inactive_and_non_clients(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "hidden.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        hidden = _seed_counterparty(conn, "00-3")
        conn.execute(
            "UPDATE counterparties SET is_active = 0, is_client = 0 WHERE id = ?",
            (hidden,),
        )
    assert _cards(db_path, "ромаш") == []


def test_admin_renames_contract_without_touching_the_date(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "rename.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        other_id = _seed_counterparty(conn, "00-2")
        _insert_contract(conn, counterparty_id=client_id, number="1028/09/26")
        _insert_contract(conn, counterparty_id=other_id, number="1029/09/26")
        contract_id = int(conn.execute("SELECT id FROM supply_contract WHERE number = '1028/09/26'").fetchone()[0])

    service = SupplyContractService(db_path=db_path)
    with pytest.raises(SupplyContractForbiddenError):
        service.update_contract(contract_id, user=MANAGER, number="1030/09/26")

    updated = service.update_contract(contract_id, user=ADMIN, number="1030/09/26")
    assert updated["number"] == "1030/09/26"
    assert updated["contract_date"] == "2026-09-25"

    with pytest.raises(SupplyContractValidationError, match="уже есть"):
        service.update_contract(contract_id, user=ADMIN, number="1029/09/26")
    with pytest.raises(SupplyContractFieldError):
        service.update_contract(contract_id, user=ADMIN, number="без-цифры")

    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        event = conn.execute(
            "SELECT old_value, new_value FROM supply_contract_event WHERE field = 'number'"
        ).fetchone()
        stored = conn.execute("SELECT number FROM supply_contract WHERE id = ?", (contract_id,)).fetchone()
    assert event["old_value"] == "1028/09/26"
    assert event["new_value"] == "1030/09/26"
    assert stored["number"] == "1030/09/26"


def test_create_for_counterparty_does_not_need_an_offer(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "create-party.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
    service = SupplyContractService(db_path=db_path)

    created = service.create_for_counterparty(client_id, _buyer(), user=MANAGER)
    assert created["counterparty_id"] == client_id
    assert created["number"] == "0001/09/26"
    assert created["manager_name"] == "manager"

    again = service.create_for_counterparty(client_id, _buyer(), user=MANAGER)
    assert again["id"] == created["id"]

    with pytest.raises(SupplyContractNotFoundError):
        service.get_for_counterparty(9999, user=MANAGER)


def test_counterparty_card_routes(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    db_path = _fresh_db(tmp_path, "http-cards.db")
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    from app.core.settings import get_settings

    get_settings.cache_clear()
    kp_db_schema._schema_ready.clear()
    patch_auth_users(
        monkeypatch,
        [
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
        ],
    )
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)

    app = create_app()
    app.dependency_overrides[get_supply_contract_service] = lambda: SupplyContractService(db_path=db_path)
    app.dependency_overrides[get_counterparties_service] = lambda: CounterpartiesService(db_path=db_path)
    client = CsrfAwareTestClient(app)

    anonymous = client.get("/api/v1/commercial/archive/counterparties", params={"q": "рома"})
    assert anonymous.status_code == 401

    created = client.post(
        f"/api/v1/commercial/archive/counterparties/{client_id}/supply-contract",
        json=_buyer().model_dump(mode="json"),
        cookies=_cookie(2, "manager", "manager"),
    )
    assert created.status_code == 200
    body = created.json()
    assert body["number"] == "0001/09/26"
    assert body["manager_name"] == "manager"

    denied = client.patch(
        f"/api/v1/commercial/archive/supply-contracts/{body['id']}",
        json={"number": "1030/09/26"},
        cookies=_cookie(2, "manager", "manager"),
    )
    assert denied.status_code == 403

    renamed = client.patch(
        f"/api/v1/commercial/archive/supply-contracts/{body['id']}",
        json={"number": "1030/09/26"},
        cookies=_cookie(1, "admin", "admin"),
    )
    assert renamed.status_code == 200
    assert renamed.json()["number"] == "1030/09/26"
    assert renamed.json()["contract_date"] == "2026-09-25"

    found = client.get(
        "/api/v1/commercial/archive/counterparties",
        params={"q": "1030/09"},
        cookies=_cookie(2, "manager", "manager"),
    )
    assert found.status_code == 200
    payload = found.json()
    assert payload["count"] == 1
    assert payload["items"][0]["id"] == client_id
    assert payload["items"][0]["contract_number"] == "1030/09/26"


def _cookie(user_id: int, username: str, role: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }
