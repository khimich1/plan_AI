"""Task T10: Generation API end-to-end (TDD).

Acceptance (D10, D6, D13):
- Generate creates draft waybills; re-generate overwrites draft only;
  confirmed/exported → 409 without force
- Start fuel/odo from last confirmed PL else from request body fields
- GET returns per-day route, km, driver, fuel_start/issued/end, odometer, warnings
- Period = request from/to (explicit period accepted by API)
- unsolvable balance → 200 with problematic_days (period is saved)
- no routes / no driver / no vehicle → 422 (configuration only)
- REQUIRE_ACCOUNTING
"""

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

from app.services.gsm_generation_service import (  # noqa: E402
    GsmGenerationError,
    GsmGenerationService,
)

PREFIX = "/api/v1/gsm"
GENERATE = f"{PREFIX}/waybills/generate"
WAYBILLS = f"{PREFIX}/waybills"

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


# =============================================================================
# Fixtures / seed helpers
# =============================================================================


def _fresh_db(tmp_path: Path, name: str = "gsm_gen.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


@pytest.fixture()
def db_path(tmp_path: Path) -> str:
    return _fresh_db(tmp_path)


@pytest.fixture()
def repo(db_path: str) -> GsmRepository:
    return GsmRepository(db_path=db_path)


@pytest.fixture()
def service(repo: GsmRepository) -> GsmGenerationService:
    return GsmGenerationService(repo=repo, holidays=frozenset(), extra_workdays=frozenset())


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


def _seed_vehicle_bundle(
    repo: GsmRepository,
    *,
    with_routes: bool = True,
    with_tx: bool = True,
    fuel_qty: float = 40.0,
    tx_day: date = date(2025, 4, 7),
) -> dict[str, int]:
    """Driver + vehicle + card + station + burn/anchor routes + one fuel tx."""
    driver_id = repo.create_driver(
        full_name="Кулигин Никита Валерьевич",
        license_number="44 21 846315",
    )
    vehicle_id = repo.create_vehicle(
        name="Geely Tugella 848",
        plate_number="О 848 ХР 44",
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    card_id = repo.create_card(
        card_number="3005454268",
        vehicle_id=vehicle_id,
        assigned_at="2025-01-01",
    )
    station_id = repo.create_station(address="АЗС Тест 1", brand="TATNEFT")

    if with_routes:
        repo.create_route(
            vehicle_id=vehicle_id,
            addr_a="Завод",
            addr_b="Объект А",
            km=190,
            frequency=100,
            typical_station_ids=json.dumps([station_id]),
        )
        for i, km in enumerate((160, 180, 200, 220), start=2):
            repo.create_route(
                vehicle_id=vehicle_id,
                addr_a="Завод",
                addr_b=f"Burn {i}",
                km=km,
                frequency=50 - i,
                typical_station_ids="[]",
            )

    if with_tx:
        batch_id = repo.create_import_batch(
            filename="seed.xls",
            uploaded_at="2025-04-30T12:00:00",
            uploaded_by="accountant",
            period_from="2025-04-01",
            period_to="2025-04-30",
        )
        repo.insert_transaction(
            card_id=card_id,
            ts=datetime(tx_day.year, tx_day.month, tx_day.day, 10, 0, 0).isoformat(
                timespec="seconds"
            ),
            service_type="fuel",
            fuel_grade="АИ-95",
            qty_liters=fuel_qty,
            amount=2500.0,
            station_id=station_id,
            raw_address="АЗС Тест 1",
            batch_id=batch_id,
        )

    repo.set_setting("winter_start", "11-01")
    repo.set_setting("hook_threshold_km", "13")

    return {
        "driver_id": driver_id,
        "vehicle_id": vehicle_id,
        "card_id": card_id,
        "station_id": station_id,
    }


def _seed_into_api_db(
    api_client: CsrfAwareTestClient,
    *,
    with_routes: bool = True,
    with_tx: bool = True,
    fuel_qty: float = 40.0,
    tx_day: date = date(2025, 4, 7),
) -> dict[str, int]:
    from app.core.settings import get_settings as _gs

    db_path = str(_gs().plita_db_path)
    repo = GsmRepository(db_path=db_path)
    return _seed_vehicle_bundle(
        repo,
        with_routes=with_routes,
        with_tx=with_tx,
        fuel_qty=fuel_qty,
        tx_day=tx_day,
    )


# =============================================================================
# Service-level
# =============================================================================


def test_serialize_waybill_route_uses_legs_or_falls_back_to_route() -> None:
    from app.services.gsm_generation_service import _serialize_waybill_route
    from core.gsm.models import LegPlan, RouteRef, TankState, WaybillDay

    route = RouteRef(route_id=7, addr_a="A", addr_b="B", km=190)
    tank = TankState(
        date=date(2025, 4, 7),
        fuel_start=20.0,
        fuel_issued=40.0,
        fuel_end=24.28,
        km=380,
        odometer_start=10_000,
        odometer_end=10_380,
    )
    with_legs = WaybillDay(
        date=date(2025, 4, 7),
        driver_id=1,
        route=route,
        tank=tank,
        source="auto",
        warnings=(),
        legs=(
            LegPlan(route_id=7, addr_a="A", addr_b="B", km=190),
            LegPlan(route_id=7, addr_a="B", addr_b="A", km=190),
        ),
    )
    parsed_legs = json.loads(_serialize_waybill_route(with_legs))
    assert parsed_legs == [
        {"from": "A", "to": "B", "km": 190, "route_id": 7},
        {"from": "B", "to": "A", "km": 190, "route_id": 7},
    ]

    no_legs = WaybillDay(
        date=date(2025, 4, 7),
        driver_id=1,
        route=route,
        tank=tank,
        source="manual",
        warnings=(),
    )
    parsed_fallback = json.loads(_serialize_waybill_route(no_legs))
    assert parsed_fallback == [{"from": "A", "to": "B", "km": 190, "route_id": 7}]


def test_service_generate_creates_draft_waybills(service: GsmGenerationService, repo: GsmRepository) -> None:
    ids = _seed_vehicle_bundle(repo)
    result = service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )
    assert len(result.waybills) >= 1
    assert all(wb.status == "draft" for wb in result.waybills)
    assert all(wb.source == "auto" for wb in result.waybills)

    listed = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    assert len(listed) == len(result.waybills)
    day = next(w for w in listed if w.date == "2025-04-07")
    assert day.driver_id == ids["driver_id"]
    assert day.fuel_issued == pytest.approx(40.0)
    assert day.fuel_start is not None
    assert day.fuel_end is not None
    assert day.odometer_start == 10_000 or day.odometer_start is not None
    assert day.km > 0
    assert day.route
    assert len(day.route) == 2
    assert day.route[0].from_addr
    assert day.route[0].to_addr
    assert day.route[1].from_addr == day.route[0].to_addr
    assert day.route[1].to_addr == day.route[0].from_addr
    assert day.km == day.route[0].km + day.route[1].km
    assert result.problematic_days == []
    assert result.manual_days == 0


def test_service_start_from_last_confirmed_else_request(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle_bundle(repo)
    repo.upsert_waybill(
        vehicle_id=ids["vehicle_id"],
        date=date(2025, 3, 31),
        driver_id=ids["driver_id"],
        status="confirmed",
        source="imported",
        odometer_start=9_800,
        odometer_end=10_000,
        fuel_start=15.0,
        fuel_issued=0.0,
        fuel_end=12.5,
        route_json="[]",
    )

    result = service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        # body fields ignored when confirmed exists
        fuel_start=99.0,
        odometer_start=1,
    )
    assert result.waybills
    first = min(result.waybills, key=lambda w: w.date)
    assert first.fuel_start == pytest.approx(12.5)
    assert first.odometer_start == 10_000


def test_service_requires_start_fields_without_confirmed(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle_bundle(repo)
    with pytest.raises(GsmGenerationError) as exc_info:
        service.generate(
            vehicle_id=ids["vehicle_id"],
            period_from=date(2025, 4, 1),
            period_to=date(2025, 4, 30),
        )
    assert exc_info.value.code == "gsm_start_required"


def test_service_regenerate_overwrites_draft_keeps_confirmed(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle_bundle(repo)
    service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )
    # Confirm the anchor day
    repo.upsert_waybill(
        vehicle_id=ids["vehicle_id"],
        date=date(2025, 4, 7),
        driver_id=ids["driver_id"],
        status="confirmed",
        source="auto",
        odometer_start=10_000,
        odometer_end=10_190,
        fuel_start=20.0,
        fuel_issued=40.0,
        fuel_end=35.0,
        route_json='[{"from":"Завод","to":"Объект А","km":190}]',
    )

    with pytest.raises(GsmGenerationError) as exc_info:
        service.generate(
            vehicle_id=ids["vehicle_id"],
            period_from=date(2025, 4, 1),
            period_to=date(2025, 4, 30),
            fuel_start=20.0,
            odometer_start=10_000,
            force=False,
        )
    assert exc_info.value.code == "gsm_confirmed_conflict"

    # force=True overwrites confirmed too
    result = service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
        force=True,
    )
    assert result.waybills
    listed = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    assert all(w.status == "draft" for w in listed)


def test_service_regenerate_overwrites_draft_only(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle_bundle(repo)
    first = service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )
    assert first.waybills
    # Re-generate with different start → drafts replaced
    second = service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=18.0,
        odometer_start=10_500,
    )
    assert second.waybills
    listed = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    assert len(listed) == len(second.waybills)
    first_day = min(listed, key=lambda w: w.date)
    assert first_day.fuel_start == pytest.approx(18.0)
    assert first_day.odometer_start == 10_500


def test_service_unsolvable_returns_problematic_days(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    """Fri→Mon overflow: period is saved, problematic_days is filled (not gsm_unsolvable)."""
    ids = _seed_vehicle_bundle(repo, with_tx=False, with_routes=True)
    batch_id = repo.create_import_batch(
        filename="dense.xls",
        uploaded_at="2025-04-30T12:00:00",
        uploaded_by="accountant",
    )
    # Proven core case: Fri+Mon 40L, start 20L, no weekday between anchors.
    for day, qty in ((date(2025, 4, 4), 40.0), (date(2025, 4, 7), 40.0)):
        repo.insert_transaction(
            card_id=ids["card_id"],
            ts=datetime(day.year, day.month, day.day, 10, 0, 0).isoformat(timespec="seconds"),
            service_type="fuel",
            qty_liters=qty,
            amount=3000.0,
            station_id=ids["station_id"],
            raw_address="АЗС Тест 1",
            batch_id=batch_id,
        )

    result = service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
        fuel_start=20.0,
        odometer_start=10_000,
    )
    assert result.waybills
    assert result.days_created == len(result.waybills)
    assert result.problematic_days
    assert result.manual_days == len(result.problematic_days)
    problem = next(p for p in result.problematic_days if p.date == "2025-04-07")
    assert problem.reason == "manual_intervention"
    assert problem.detail
    assert problem.fuel_to_issue == pytest.approx(40.0)
    assert problem.tank_volume == pytest.approx(55.0)
    listed = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 1),
        period_to=date(2025, 4, 30),
    )
    assert len(listed) == result.days_created
    monday = next(w for w in listed if w.date == "2025-04-07")
    assert "manual_intervention" in monday.warnings


def test_service_no_routes_raises(service: GsmGenerationService, repo: GsmRepository) -> None:
    """Transactions without a route library remain a configuration error."""
    ids = _seed_vehicle_bundle(repo, with_routes=False, with_tx=True)
    with pytest.raises(GsmGenerationError) as exc_info:
        service.generate(
            vehicle_id=ids["vehicle_id"],
            period_from=date(2025, 4, 1),
            period_to=date(2025, 4, 30),
            fuel_start=20.0,
            odometer_start=10_000,
        )
    assert exc_info.value.code == "gsm_routes_required"


# =============================================================================
# HTTP-level
# =============================================================================


def test_http_generate_and_list(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth()

    gen = api_client.post(
        GENERATE,
        json={
            "vehicle_id": ids["vehicle_id"],
            "period_from": "2025-04-01",
            "period_to": "2025-04-30",
            "fuel_start": 20.0,
            "odometer_start": 10000,
        },
        cookies=cookies,
    )
    assert gen.status_code == 200, gen.text
    body = gen.json()
    assert "waybills" in body
    assert len(body["waybills"]) >= 1
    assert body["waybills"][0]["status"] == "draft"
    assert "warnings" in body
    assert body["problematic_days"] == []
    assert body["manual_days"] == 0

    listed = api_client.get(
        WAYBILLS,
        params={
            "vehicle_id": ids["vehicle_id"],
            "from": "2025-04-01",
            "to": "2025-04-30",
        },
        cookies=cookies,
    )
    assert listed.status_code == 200, listed.text
    days = listed.json()
    assert isinstance(days, list)
    assert len(days) >= 1
    sample = days[0]
    for key in (
        "date",
        "driver_id",
        "km",
        "fuel_start",
        "fuel_issued",
        "fuel_end",
        "odometer_start",
        "odometer_end",
        "route",
        "warnings",
        "status",
    ):
        assert key in sample


def test_http_generate_409_without_force(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth()
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))

    assert (
        api_client.post(
            GENERATE,
            json={
                "vehicle_id": ids["vehicle_id"],
                "period_from": "2025-04-01",
                "period_to": "2025-04-30",
                "fuel_start": 20.0,
                "odometer_start": 10000,
            },
            cookies=cookies,
        ).status_code
        == 200
    )

    repo.upsert_waybill(
        vehicle_id=ids["vehicle_id"],
        date=date(2025, 4, 7),
        driver_id=ids["driver_id"],
        status="exported",
        source="auto",
        odometer_start=10000,
        odometer_end=10190,
        fuel_start=20.0,
        fuel_issued=40.0,
        fuel_end=35.0,
        route_json="[]",
    )

    conflict = api_client.post(
        GENERATE,
        json={
            "vehicle_id": ids["vehicle_id"],
            "period_from": "2025-04-01",
            "period_to": "2025-04-30",
            "fuel_start": 20.0,
            "odometer_start": 10000,
        },
        cookies=cookies,
    )
    assert conflict.status_code == 409
    detail = conflict.json()["detail"]
    assert detail["code"] == "gsm_confirmed_conflict"

    forced = api_client.post(
        GENERATE,
        json={
            "vehicle_id": ids["vehicle_id"],
            "period_from": "2025-04-01",
            "period_to": "2025-04-30",
            "fuel_start": 20.0,
            "odometer_start": 10000,
            "force": True,
        },
        cookies=cookies,
    )
    assert forced.status_code == 200, forced.text


def test_http_unsolvable_200_with_problematic_days(api_client: CsrfAwareTestClient) -> None:
    """Nonsolvable Fri→Mon anchor → 200 with problematic_days, not 422 gsm_unsolvable."""
    ids = _seed_into_api_db(api_client, with_tx=False)
    cookies = _auth()
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    batch_id = repo.create_import_batch(
        filename="dense.xls",
        uploaded_at="2025-04-30T12:00:00",
        uploaded_by="accountant",
    )
    for day, qty in ((date(2025, 4, 4), 40.0), (date(2025, 4, 7), 40.0)):
        repo.insert_transaction(
            card_id=ids["card_id"],
            ts=datetime(day.year, day.month, day.day, 10, 0, 0).isoformat(timespec="seconds"),
            service_type="fuel",
            qty_liters=qty,
            amount=3000.0,
            station_id=ids["station_id"],
            raw_address="АЗС Тест 1",
            batch_id=batch_id,
        )

    resp = api_client.post(
        GENERATE,
        json={
            "vehicle_id": ids["vehicle_id"],
            "period_from": "2025-04-01",
            "period_to": "2025-04-30",
            "fuel_start": 20.0,
            "odometer_start": 10000,
        },
        cookies=cookies,
    )
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["days_created"] >= 1
    assert body["waybills"]
    problems = body["problematic_days"]
    assert problems
    assert body["manual_days"] == len(problems)
    monday = next(p for p in problems if p["date"] == "2025-04-07")
    assert monday["reason"] == "manual_intervention"
    assert monday["detail"]
    assert monday["fuel_before"] > 0
    assert monday["fuel_to_issue"] == pytest.approx(40.0)
    assert monday["tank_volume"] == pytest.approx(55.0)


def test_http_no_routes_422(api_client: CsrfAwareTestClient) -> None:
    """Transactions without a route library stay 422 gsm_routes_required."""
    ids = _seed_into_api_db(api_client, with_routes=False)
    cookies = _auth()

    resp = api_client.post(
        GENERATE,
        json={
            "vehicle_id": ids["vehicle_id"],
            "period_from": "2025-04-01",
            "period_to": "2025-04-30",
            "fuel_start": 20.0,
            "odometer_start": 10000,
        },
        cookies=cookies,
    )
    assert resp.status_code == 422
    assert resp.json()["detail"]["code"] == "gsm_routes_required"


def test_http_manager_forbidden(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth("manager_a")

    gen = api_client.post(
        GENERATE,
        json={
            "vehicle_id": ids["vehicle_id"],
            "period_from": "2025-04-01",
            "period_to": "2025-04-30",
            "fuel_start": 20.0,
            "odometer_start": 10000,
        },
        cookies=cookies,
    )
    assert gen.status_code == 403

    listed = api_client.get(
        WAYBILLS,
        params={"vehicle_id": ids["vehicle_id"], "from": "2025-04-01", "to": "2025-04-30"},
        cookies=cookies,
    )
    assert listed.status_code == 403


def test_http_accountant_and_admin_allowed(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    for user in ("accountant_user", "admin"):
        resp = api_client.post(
            GENERATE,
            json={
                "vehicle_id": ids["vehicle_id"],
                "period_from": "2025-04-01",
                "period_to": "2025-04-30",
                "fuel_start": 20.0,
                "odometer_start": 10000,
                "force": True,
            },
            cookies=_auth(user),
        )
        assert resp.status_code == 200, f"{user}: {resp.text}"
