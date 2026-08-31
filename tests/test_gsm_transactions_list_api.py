"""GET /api/v1/gsm/transactions — fleet journal (T2)."""

from __future__ import annotations

from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.gsm_repository import GsmRepository
from core import kp_db_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

PREFIX = "/api/v1/gsm"
TX_URL = f"{PREFIX}/transactions"

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
        "id": 5,
        "username": "accountant_user",
        "role": "accountant",
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
    {
        "id": 2,
        "username": "prod_user",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


@pytest.fixture()
def api_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> CsrfAwareTestClient:
    db = tmp_path / "plita.db"
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db))
    monkeypatch.setenv("PB_DB_PATH", str(db))
    get_settings.cache_clear()
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(str(db))
    patch_auth_users(monkeypatch, TEST_USERS)
    return CsrfAwareTestClient(create_app())


def _auth(username: str = "accountant_user") -> dict[str, str]:
    by_user = {
        "admin": (1, "admin"),
        "accountant_user": (5, "accountant"),
        "manager_a": (3, "manager"),
        "prod_user": (2, "production"),
    }
    user_id, role = by_user[username]
    return session_cookie(user_id, role, username)


def _seed_fleet(api_client: CsrfAwareTestClient) -> dict[str, int]:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 846315")
    v1 = repo.create_vehicle(
        name="Car 1",
        plate_number="A111AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    v2 = repo.create_vehicle(
        name="Car 2",
        plate_number="B222BB44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    c1 = repo.create_card(card_number="111", vehicle_id=v1, assigned_at="2026-01-01")
    c2 = repo.create_card(card_number="222", vehicle_id=v2, assigned_at="2026-01-01")
    unbound = repo.create_card(card_number="orphan", vehicle_id=None, assigned_at="2026-01-01")
    station_id = repo.create_station(address="АЗС 1", brand="TATNEFT")
    batch_id = repo.create_import_batch(
        filename="fleet.xls",
        uploaded_at="2026-08-14T12:00:00",
    )
    repo.insert_transaction(
        card_id=c1,
        ts="2026-08-03T10:00:00",
        service_type="fuel",
        fuel_grade="АИ-95",
        qty_liters=40.0,
        amount=2500.0,
        station_id=station_id,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )
    repo.insert_transaction(
        card_id=c1,
        ts="2026-08-10T12:00:00",
        service_type="wash",
        qty_liters=None,
        amount=500.0,
        raw_address="Мойка 1",
        batch_id=batch_id,
    )
    repo.insert_transaction(
        card_id=c2,
        ts="2026-08-05T09:00:00",
        service_type="fuel",
        fuel_grade="АИ-95",
        qty_liters=10.0,
        amount=700.0,
        station_id=station_id,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )
    repo.insert_transaction(
        card_id=unbound,
        ts="2026-08-11T08:00:00",
        service_type="fuel",
        fuel_grade="АИ-95",
        qty_liters=50.0,
        amount=3000.0,
        raw_address="АЗС 2",
        batch_id=batch_id,
    )
    # outside period
    repo.insert_transaction(
        card_id=c1,
        ts="2026-07-31T10:00:00",
        service_type="fuel",
        qty_liters=99.0,
        amount=1.0,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )
    return {"v1": v1, "v2": v2, "station_id": station_id}


def _list(
    client: CsrfAwareTestClient,
    *,
    params: dict[str, object] | None = None,
    username: str = "accountant_user",
):
    return client.get(TX_URL, params=params or {}, cookies=_auth(username))


def test_list_transactions_fleet_for_period(api_client: CsrfAwareTestClient) -> None:
    _seed_fleet(api_client)
    resp = _list(api_client, params={"from": "2026-08-01", "to": "2026-08-31"})
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["total_count"] == 4
    assert body["sum_liters"] == pytest.approx(100.0)
    assert body["sum_amount"] == pytest.approx(6700.0)
    assert len(body["rows"]) == 4
    liters = sum(row["qty_liters"] or 0 for row in body["rows"])
    amounts = sum(row["amount"] for row in body["rows"])
    assert liters == pytest.approx(body["sum_liters"])
    assert amounts == pytest.approx(body["sum_amount"])
    timestamps = [row["ts"] for row in body["rows"]]
    assert timestamps == sorted(timestamps)


def test_list_transactions_filters_vehicle_service_type_period(
    api_client: CsrfAwareTestClient,
) -> None:
    ids = _seed_fleet(api_client)

    by_vehicle = _list(
        api_client,
        params={"from": "2026-08-01", "to": "2026-08-31", "vehicle_id": ids["v1"]},
    )
    assert by_vehicle.status_code == 200
    body = by_vehicle.json()
    assert body["total_count"] == 2
    assert all(row["vehicle_id"] == ids["v1"] for row in body["rows"])
    assert body["sum_liters"] == pytest.approx(40.0)
    assert body["sum_amount"] == pytest.approx(3000.0)

    by_type = _list(
        api_client,
        params={"from": "2026-08-01", "to": "2026-08-31", "service_type": "wash"},
    )
    assert by_type.json()["total_count"] == 1
    assert by_type.json()["rows"][0]["service_type"] == "wash"
    assert by_type.json()["sum_liters"] == pytest.approx(0.0)
    assert by_type.json()["sum_amount"] == pytest.approx(500.0)

    combined = _list(
        api_client,
        params={
            "from": "2026-08-01",
            "to": "2026-08-31",
            "vehicle_id": ids["v1"],
            "service_type": "fuel",
        },
    )
    assert combined.json()["total_count"] == 1
    assert combined.json()["rows"][0]["qty_liters"] == pytest.approx(40.0)

    narrow = _list(
        api_client,
        params={"from": "2026-08-05", "to": "2026-08-05"},
    )
    assert narrow.json()["total_count"] == 1
    assert narrow.json()["rows"][0]["vehicle_id"] == ids["v2"]


def test_list_transactions_empty_totals_are_zero(api_client: CsrfAwareTestClient) -> None:
    _seed_fleet(api_client)
    resp = _list(api_client, params={"from": "2026-01-01", "to": "2026-01-31"})
    assert resp.status_code == 200
    body = resp.json()
    assert body["rows"] == []
    assert body["total_count"] == 0
    assert body["sum_liters"] == pytest.approx(0.0)
    assert body["sum_amount"] == pytest.approx(0.0)


def test_list_transactions_unbound_card_has_null_vehicle(
    api_client: CsrfAwareTestClient,
) -> None:
    _seed_fleet(api_client)
    resp = _list(api_client, params={"from": "2026-08-01", "to": "2026-08-31"})
    orphan = next(row for row in resp.json()["rows"] if row["card_number"] == "orphan")
    assert orphan["vehicle_id"] is None
    assert orphan["address"] == "АЗС 2"


def test_list_transactions_forbidden_for_manager_and_production(
    api_client: CsrfAwareTestClient,
) -> None:
    params = {"from": "2026-08-01", "to": "2026-08-31"}
    for user in ("manager_a", "prod_user"):
        resp = _list(api_client, params=params, username=user)
        assert resp.status_code == 403, user
