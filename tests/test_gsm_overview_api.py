"""GET /api/v1/gsm/overview + vehicle period status model (T3)."""

from __future__ import annotations

from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.gsm_repository import GsmRepository
from app.services.gsm_overview_service import _chain_broken, _status_of
from core import kp_db_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

OVERVIEW = "/api/v1/gsm/overview"

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
]


def _agg(**overrides: object) -> dict[str, object]:
    base: dict[str, object] = {
        "tx_count": 0,
        "tx_last_date": None,
        "wb_count": 0,
        "wb_last_date": None,
        "red_days": 0,
        "draft_count": 0,
        "exported_count": 0,
    }
    base.update(overrides)
    return base


def test_status_no_data_when_empty() -> None:
    assert _status_of(_agg()) == "no_data"


def test_status_needs_generation_when_tx_without_waybills() -> None:
    assert _status_of(_agg(tx_count=3, tx_last_date="2026-08-20")) == "needs_generation"


def test_status_needs_generation_when_tx_after_last_waybill() -> None:
    assert (
        _status_of(
            _agg(
                tx_count=2,
                tx_last_date="2026-08-21",
                wb_count=2,
                wb_last_date="2026-08-20",
                draft_count=1,
            )
        )
        == "needs_generation"
    )


def test_status_has_red_days_before_needs_generation() -> None:
    assert (
        _status_of(
            _agg(
                tx_count=2,
                tx_last_date="2026-08-21",
                wb_count=2,
                wb_last_date="2026-08-20",
                red_days=1,
                draft_count=1,
            )
        )
        == "has_red_days"
    )


def test_status_equal_last_dates_is_not_needs_generation() -> None:
    assert (
        _status_of(
            _agg(
                tx_count=2,
                tx_last_date="2026-08-20",
                wb_count=2,
                wb_last_date="2026-08-20",
                exported_count=2,
            )
        )
        == "ready"
    )


def test_status_has_red_days_before_drafts() -> None:
    assert (
        _status_of(
            _agg(
                tx_count=1,
                tx_last_date="2026-08-10",
                wb_count=3,
                wb_last_date="2026-08-10",
                red_days=1,
                draft_count=2,
            )
        )
        == "has_red_days"
    )


def test_status_drafts_pending() -> None:
    assert (
        _status_of(
            _agg(
                tx_count=1,
                tx_last_date="2026-08-10",
                wb_count=2,
                wb_last_date="2026-08-10",
                draft_count=1,
                exported_count=1,
            )
        )
        == "drafts_pending"
    )


def test_status_pending_export() -> None:
    assert (
        _status_of(
            _agg(
                tx_count=1,
                tx_last_date="2026-08-10",
                wb_count=2,
                wb_last_date="2026-08-10",
                exported_count=1,
            )
        )
        == "pending_export"
    )


def test_status_ready_when_all_exported() -> None:
    assert (
        _status_of(
            _agg(
                tx_count=1,
                tx_last_date="2026-08-10",
                wb_count=2,
                wb_last_date="2026-08-10",
                exported_count=2,
            )
        )
        == "ready"
    )


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
    }
    user_id, role = by_user[username]
    return session_cookie(user_id, role, username)


def test_overview_tx_without_waybills_needs_generation(
    api_client: CsrfAwareTestClient,
) -> None:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 111111")
    vehicle_id = repo.create_vehicle(
        name="Car 1",
        plate_number="A111AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    card_id = repo.create_card(
        card_number="111", vehicle_id=vehicle_id, assigned_at="2026-01-01"
    )
    batch_id = repo.create_import_batch(
        filename="tx.xls", uploaded_at="2026-08-14T12:00:00"
    )
    repo.insert_transaction(
        card_id=card_id,
        ts="2026-08-10T10:00:00",
        service_type="fuel",
        qty_liters=40.0,
        amount=2500.0,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    rows = resp.json()
    assert len(rows) == 1
    row = rows[0]
    assert row["vehicle"]["id"] == vehicle_id
    assert row["status"] == "needs_generation"
    assert row["tx_count"] == 1
    assert row["wb_count"] == 0
    assert row["liters_diff"] == pytest.approx(40.0)
    assert row["open_before"] == 0
    assert row["open_before_month"] is None
    assert row["chain_broken"] is False
    assert set(row["vehicle"]) == {"id", "name", "plate_number"}
    assert {
        "vehicle",
        "tx_count",
        "tx_liters",
        "tx_amount",
        "tx_last_date",
        "wb_count",
        "wb_km",
        "wb_fuel_issued",
        "wb_last_date",
        "red_days",
        "draft_count",
        "confirmed_count",
        "exported_count",
        "fuel_end_last",
        "liters_diff",
        "open_before",
        "status",
    }.issubset(row.keys())
    assert "open_before_month" in row
    assert "chain_broken" in row


def test_overview_july_draft_open_before_month(
    api_client: CsrfAwareTestClient,
) -> None:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 333333")
    vehicle_id = repo.create_vehicle(
        name="Car Tail",
        plate_number="A333AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-07-20",
        driver_id=driver_id,
        odometer_start=62000,
        odometer_end=62100,
        fuel_start=20.0,
        fuel_end=15.0,
        status="draft",
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    row = resp.json()[0]
    assert row["open_before"] == 1
    assert row["open_before_month"] == "2026-07"
    assert row["chain_broken"] is False


def test_overview_chain_broken_when_july_end_mismatches_august_start(
    api_client: CsrfAwareTestClient,
) -> None:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 444444")
    vehicle_id = repo.create_vehicle(
        name="Car Chain",
        plate_number="A444AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-07-31",
        driver_id=driver_id,
        odometer_start=62700,
        odometer_end=62846,
        fuel_start=30.0,
        fuel_end=27.59,
        status="exported",
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-08-01",
        driver_id=driver_id,
        odometer_start=62946,
        odometer_end=63100,
        fuel_start=18.09,
        fuel_end=10.0,
        status="draft",
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    row = resp.json()[0]
    assert row["chain_broken"] is True
    assert row["open_before"] == 0
    assert row["open_before_month"] is None


def test_overview_chain_ok_when_tank_and_odometer_match(
    api_client: CsrfAwareTestClient,
) -> None:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 555555")
    vehicle_id = repo.create_vehicle(
        name="Car Match",
        plate_number="A555AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-07-31",
        driver_id=driver_id,
        odometer_start=62700,
        odometer_end=62846,
        fuel_start=30.0,
        fuel_end=27.59,
        status="exported",
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-08-01",
        driver_id=driver_id,
        odometer_start=62846,
        odometer_end=63100,
        fuel_start=27.59,
        fuel_end=10.0,
        status="draft",
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    row = resp.json()[0]
    assert row["chain_broken"] is False
    assert row["open_before"] == 0
    assert row["open_before_month"] is None


def test_overview_chain_not_broken_when_no_waybill_in_period(
    api_client: CsrfAwareTestClient,
) -> None:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 666666")
    vehicle_id = repo.create_vehicle(
        name="Car Gap",
        plate_number="A666AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-07-31",
        driver_id=driver_id,
        odometer_start=62700,
        odometer_end=62846,
        fuel_start=30.0,
        fuel_end=27.59,
        status="exported",
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    row = resp.json()[0]
    assert row["wb_count"] == 0
    assert row["chain_broken"] is False


def test_overview_chain_broken_when_only_odometer_mismatches(
    api_client: CsrfAwareTestClient,
) -> None:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 777777")
    vehicle_id = repo.create_vehicle(
        name="Car Odo",
        plate_number="A777AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-07-31",
        driver_id=driver_id,
        odometer_start=62700,
        odometer_end=62846,
        fuel_start=30.0,
        fuel_end=27.59,
        status="exported",
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-08-01",
        driver_id=driver_id,
        odometer_start=62946,
        odometer_end=63100,
        fuel_start=27.59,
        fuel_end=10.0,
        status="draft",
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    row = resp.json()[0]
    assert row["chain_broken"] is True
    assert row["open_before"] == 0
    assert row["open_before_month"] is None


def test_chain_broken_helper_missing_fuel_is_false() -> None:
    assert (
        _chain_broken(
            {
                "chain_prev_fuel_end": None,
                "chain_prev_odometer_end": 100,
                "chain_first_fuel_start": 10.0,
                "chain_first_odometer_start": 100,
            }
        )
        is False
    )


@pytest.mark.parametrize(
    ("overrides", "expected"),
    [
        ({}, False),
        ({"chain_first_odometer_start": 101}, True),
        ({"chain_first_fuel_start": 18.09}, True),
        ({"chain_first_fuel_start": 10.01}, False),
        ({"chain_first_fuel_start": 10.011}, True),
    ],
)
def test_overview_chain_broken_helper_odo_or_fuel_mismatch(
    overrides: dict[str, object],
    expected: bool,
) -> None:
    row: dict[str, object] = {
        "chain_prev_fuel_end": 10.0,
        "chain_prev_odometer_end": 100,
        "chain_first_fuel_start": 10.0,
        "chain_first_odometer_start": 100,
    }
    row.update(overrides)
    assert _chain_broken(row) is expected


def test_overview_forbidden_for_manager(api_client: CsrfAwareTestClient) -> None:
    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth("manager_a"),
    )
    assert resp.status_code == 403


def test_overview_invalid_period_400(api_client: CsrfAwareTestClient) -> None:
    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-31", "to": "2026-08-01"},
        cookies=_auth(),
    )
    assert resp.status_code == 400
    assert resp.json()["detail"]["code"] == "gsm_invalid_period"


def test_overview_tx_liters_excludes_wash_qty(
    api_client: CsrfAwareTestClient,
) -> None:
    """Wash qty_liters must not enter tx_liters or liters_diff (fuel vs issued)."""
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 222222")
    vehicle_id = repo.create_vehicle(
        name="Car Wash",
        plate_number="A222AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    card_id = repo.create_card(
        card_number="222", vehicle_id=vehicle_id, assigned_at="2026-01-01"
    )
    batch_id = repo.create_import_batch(
        filename="tx.xls", uploaded_at="2026-08-14T12:00:00"
    )
    repo.insert_transaction(
        card_id=card_id,
        ts="2026-08-10T10:00:00",
        service_type="fuel",
        qty_liters=40.0,
        amount=2500.0,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )
    repo.insert_transaction(
        card_id=card_id,
        ts="2026-08-15T12:00:00",
        service_type="wash",
        qty_liters=1.0,
        amount=500.0,
        raw_address="Мойка 1",
        batch_id=batch_id,
    )
    repo.insert_transaction(
        card_id=card_id,
        ts="2026-08-20T11:00:00",
        service_type="fuel",
        qty_liters=25.0,
        amount=1600.0,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-08-20",
        driver_id=driver_id,
        odometer_start=10000,
        odometer_end=10200,
        fuel_issued=65.0,
        fuel_end=30.0,
        status="exported",
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    rows = resp.json()
    assert len(rows) == 1
    row = rows[0]
    fuel_liters = 65.0
    assert row["tx_count"] == 3
    assert row["tx_liters"] == pytest.approx(fuel_liters)
    assert row["wb_fuel_issued"] == pytest.approx(fuel_liters)
    assert row["liters_diff"] == pytest.approx(0.0)
    assert row["red_days"] == 0
    assert row["status"] == "ready"


def test_overview_red_days_beats_newer_transaction(
    api_client: CsrfAwareTestClient,
) -> None:
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    driver_id = repo.create_driver(full_name="Driver", license_number="44 21 666666")
    vehicle_id = repo.create_vehicle(
        name="Car Red",
        plate_number="A666AA44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    card_id = repo.create_card(
        card_number="666", vehicle_id=vehicle_id, assigned_at="2026-01-01"
    )
    batch_id = repo.create_import_batch(
        filename="tx.xls", uploaded_at="2026-08-22T12:00:00"
    )
    repo.upsert_waybill(
        vehicle_id=vehicle_id,
        date="2026-08-10",
        driver_id=driver_id,
        odometer_start=10000,
        odometer_end=10100,
        fuel_start=20.0,
        fuel_issued=10.0,
        fuel_end=15.0,
        status="draft",
        warnings_json='["manual_intervention"]',
    )
    repo.insert_transaction(
        card_id=card_id,
        ts="2026-08-21T10:00:00",
        service_type="fuel",
        qty_liters=30.0,
        amount=2000.0,
        raw_address="АЗС 1",
        batch_id=batch_id,
    )

    resp = api_client.get(
        OVERVIEW,
        params={"from": "2026-08-01", "to": "2026-08-31"},
        cookies=_auth(),
    )
    assert resp.status_code == 200, resp.text
    row = resp.json()[0]
    assert row["vehicle"]["id"] == vehicle_id
    assert row["red_days"] == 1
    assert row["tx_last_date"] == "2026-08-21"
    assert row["wb_last_date"] == "2026-08-10"
    assert row["status"] == "has_red_days"
