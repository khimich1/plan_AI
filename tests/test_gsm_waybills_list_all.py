"""GET /api/v1/gsm/waybills — vehicle_id optional (T4)."""

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

WAYBILLS = "/api/v1/gsm/waybills"

TEST_USERS = [
    {
        "id": 5,
        "username": "accountant_user",
        "role": "accountant",
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


def _auth() -> dict[str, str]:
    return session_cookie(5, "accountant", "accountant_user")


def _seed_two_vehicles(api_client: CsrfAwareTestClient) -> dict[str, int]:
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
    repo.upsert_waybill(
        vehicle_id=v2,
        date="2026-08-01",
        driver_id=driver_id,
        odometer_start=1,
        odometer_end=10,
        route_json='[{"from":"A","to":"B","km":9}]',
    )
    repo.upsert_waybill(
        vehicle_id=v1,
        date="2026-08-10",
        driver_id=driver_id,
        odometer_start=1,
        odometer_end=5,
        route_json='[{"from":"A","to":"C","km":4}]',
    )
    repo.upsert_waybill(
        vehicle_id=v1,
        date="2026-08-02",
        driver_id=driver_id,
        odometer_start=1,
        odometer_end=3,
        route_json='[{"from":"A","to":"D","km":2}]',
    )
    return {"v1": v1, "v2": v2}


def test_list_waybills_all_vehicles_sorted(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_two_vehicles(api_client)
    resp = api_client.get(
        WAYBILLS,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    rows = resp.json()
    assert [(row["vehicle_id"], row["date"]) for row in rows] == [
        (ids["v1"], "2026-08-02"),
        (ids["v1"], "2026-08-10"),
        (ids["v2"], "2026-08-01"),
    ]


def test_list_waybills_with_vehicle_id_unchanged(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_two_vehicles(api_client)
    resp = api_client.get(
        WAYBILLS,
        params={"vehicle_id": ids["v2"], "from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    rows = resp.json()
    assert len(rows) == 1
    assert rows[0]["vehicle_id"] == ids["v2"]
    assert rows[0]["date"] == "2026-08-01"
