"""POST /api/v1/gsm/waybills/generate-bulk (T5)."""

from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.gsm_repository import GsmRepository
from core import kp_db_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

BULK = "/api/v1/gsm/waybills/generate-bulk"

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


def _repo() -> GsmRepository:
    from app.core.settings import get_settings as _gs

    return GsmRepository(db_path=str(_gs().plita_db_path))


def _seed_vehicle(
    repo: GsmRepository,
    *,
    name: str,
    plate: str,
    card_number: str,
    with_routes: bool,
    with_tx: bool,
    with_confirmed_start: bool,
) -> dict[str, int]:
    driver_id = repo.create_driver(full_name=f"Drv {name}", license_number=f"44 {plate[:2]} 111111")
    vehicle_id = repo.create_vehicle(
        name=name,
        plate_number=plate,
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    card_id = repo.create_card(
        card_number=card_number, vehicle_id=vehicle_id, assigned_at="2025-01-01"
    )
    station = repo.get_or_create_station(address="АЗС Тест 1", brand="TATNEFT")
    if with_routes:
        repo.create_route(
            vehicle_id=vehicle_id,
            addr_a="Завод",
            addr_b="Объект А",
            km=190,
            frequency=100,
            typical_station_ids=json.dumps([station]),
        )
        for i, km in enumerate((160, 180, 200, 220), start=2):
            repo.create_route(
                vehicle_id=vehicle_id,
                addr_a="Завод",
                addr_b=f"Burn {name} {i}",
                km=km,
                frequency=50 - i,
                typical_station_ids="[]",
            )
    if with_tx:
        batch_id = repo.create_import_batch(
            filename=f"{name}.xls",
            uploaded_at="2025-04-30T12:00:00",
        )
        repo.insert_transaction(
            card_id=card_id,
            ts=datetime(2025, 4, 7, 10, 0, 0).isoformat(timespec="seconds"),
            service_type="fuel",
            fuel_grade="АИ-95",
            qty_liters=40.0,
            amount=2500.0,
            station_id=station,
            raw_address="АЗС Тест 1",
            batch_id=batch_id,
        )
    if with_confirmed_start:
        repo.upsert_waybill(
            vehicle_id=vehicle_id,
            date=date(2025, 3, 31),
            driver_id=driver_id,
            status="exported",
            source="imported",
            odometer_start=9_800,
            odometer_end=10_000,
            fuel_start=15.0,
            fuel_issued=0.0,
            fuel_end=12.5,
            route_json="[]",
        )
    repo.set_setting("winter_start", "11-01")
    repo.set_setting("hook_threshold_km", "13")
    return {"driver_id": driver_id, "vehicle_id": vehicle_id}


def _bulk(
    client: CsrfAwareTestClient,
    *,
    vehicle_ids: list[int],
    period_from: str = "2025-04-01",
    period_to: str = "2025-04-30",
    force: bool = False,
):
    return client.post(
        BULK,
        json={
            "vehicle_ids": vehicle_ids,
            "period_from": period_from,
            "period_to": period_to,
            "force": force,
        },
        cookies=_auth(),
    )


def test_bulk_fleet_empty_routes_required(api_client: CsrfAwareTestClient) -> None:
    """gsm_routes_required fires when the fleet has zero routes."""
    repo = _repo()
    empty_a = _seed_vehicle(
        repo,
        name="EmptyA",
        plate="A111AA44",
        card_number="111",
        with_routes=False,
        with_tx=True,
        with_confirmed_start=True,
    )
    empty_b = _seed_vehicle(
        repo,
        name="EmptyB",
        plate="B222BB44",
        card_number="222",
        with_routes=False,
        with_tx=True,
        with_confirmed_start=True,
    )
    resp = _bulk(
        api_client, vehicle_ids=[empty_a["vehicle_id"], empty_b["vehicle_id"]]
    )
    assert resp.status_code == 200, resp.text
    results = {item["vehicle_id"]: item for item in resp.json()["results"]}
    assert results[empty_a["vehicle_id"]]["ok"] is False
    assert results[empty_a["vehicle_id"]]["error"]["code"] == "gsm_routes_required"
    assert results[empty_b["vehicle_id"]]["ok"] is False
    assert results[empty_b["vehicle_id"]]["error"]["code"] == "gsm_routes_required"


def test_bulk_sibling_routes_allow_generate(api_client: CsrfAwareTestClient) -> None:
    """A vehicle with no own routes generates from sibling fleet routes."""
    repo = _repo()
    donor = _seed_vehicle(
        repo,
        name="Ok",
        plate="A111AA44",
        card_number="111",
        with_routes=True,
        with_tx=True,
        with_confirmed_start=True,
    )
    borrower = _seed_vehicle(
        repo,
        name="Bad",
        plate="B222BB44",
        card_number="222",
        with_routes=False,
        with_tx=True,
        with_confirmed_start=True,
    )
    resp = _bulk(
        api_client, vehicle_ids=[borrower["vehicle_id"], donor["vehicle_id"]]
    )
    assert resp.status_code == 200, resp.text
    results = {item["vehicle_id"]: item for item in resp.json()["results"]}
    assert results[donor["vehicle_id"]]["ok"] is True
    assert results[donor["vehicle_id"]]["result"]["days_created"] >= 1
    assert results[borrower["vehicle_id"]]["ok"] is True
    borrower_result = results[borrower["vehicle_id"]]["result"]
    assert borrower_result["days_created"] >= 1
    borrower_legs = [
        leg
        for waybill in borrower_result["waybills"]
        for leg in waybill.get("route") or []
    ]
    assert any(int(leg.get("km") or 0) > 0 for leg in borrower_legs)
    # Sibling library route: persist addresses/km, not the donor's route_id.
    assert all(
        leg.get("route_id") is None
        for leg in borrower_legs
        if int(leg.get("km") or 0) > 0
    )


def test_bulk_start_required_is_per_vehicle_not_http(
    api_client: CsrfAwareTestClient,
) -> None:
    repo = _repo()
    ids = _seed_vehicle(
        repo,
        name="NoStart",
        plate="C333CC44",
        card_number="333",
        with_routes=True,
        with_tx=True,
        with_confirmed_start=False,
    )
    resp = _bulk(api_client, vehicle_ids=[ids["vehicle_id"]])
    assert resp.status_code == 200, resp.text
    item = resp.json()["results"][0]
    assert item["ok"] is False
    assert item["error"]["code"] == "gsm_start_required"


def test_bulk_force_overwrites_confirmed(api_client: CsrfAwareTestClient) -> None:
    repo = _repo()
    ids = _seed_vehicle(
        repo,
        name="Force",
        plate="D444DD44",
        card_number="444",
        with_routes=True,
        with_tx=True,
        with_confirmed_start=True,
    )
    first = _bulk(api_client, vehicle_ids=[ids["vehicle_id"]])
    assert first.status_code == 200
    assert first.json()["results"][0]["ok"] is True

    repo.upsert_waybill(
        vehicle_id=ids["vehicle_id"],
        date=date(2025, 4, 7),
        driver_id=ids["driver_id"],
        status="confirmed",
        source="auto",
        odometer_start=10_000,
        odometer_end=10_190,
        fuel_start=12.5,
        fuel_issued=40.0,
        fuel_end=35.0,
        route_json="[]",
    )
    blocked = _bulk(api_client, vehicle_ids=[ids["vehicle_id"]], force=False)
    assert blocked.status_code == 200
    assert blocked.json()["results"][0]["error"]["code"] == "gsm_confirmed_conflict"

    forced = _bulk(api_client, vehicle_ids=[ids["vehicle_id"]], force=True)
    assert forced.status_code == 200, forced.text
    assert forced.json()["results"][0]["ok"] is True


def test_bulk_empty_vehicle_ids(api_client: CsrfAwareTestClient) -> None:
    resp = _bulk(api_client, vehicle_ids=[])
    assert resp.status_code == 200, resp.text
    assert resp.json()["results"] == []


def test_bulk_invalid_period_400(api_client: CsrfAwareTestClient) -> None:
    resp = _bulk(
        api_client,
        vehicle_ids=[1],
        period_from="2025-04-30",
        period_to="2025-04-01",
    )
    assert resp.status_code == 400
    assert resp.json()["detail"]["code"] == "gsm_invalid_period"
