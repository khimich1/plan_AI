"""Season mode (manual switch) + waybill immutability / cascade re-chain.

D6b Season is switched MANUALLY (accountant button), not by calendar:
    ``season_switches`` setting = [{"date": "YYYY-MM-DD", "mode": ...}];
    mode(day) = mode of the latest switch ≤ day; none → summer.
D11b PATCH /gsm/waybills/{id}: confirmed/exported waybill → 409
    ``gsm_waybill_locked``; any confirmed/exported day AFTER the edited one →
    409 ``gsm_chain_locked`` (nothing persisted); otherwise downstream drafts
    are rechained and the response reports ``rechained_draft_days``.
"""

from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.gsm_repository import GsmRepository
from app.services.gsm_generation_service import GsmGenerationError, GsmGenerationService
from app.services.gsm_registry_service import GsmRegistryError, GsmRegistryService
from core import kp_db_schema
from core.gsm.balance import burn_for_km
from core.gsm.generator import LibraryRoute, generate
from core.gsm.models import Transaction
from core.gsm.season import norm_for, season_mode_for
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

PREFIX = "/api/v1/gsm"
SETTINGS = f"{PREFIX}/settings"
SEASON = f"{PREFIX}/settings/season"
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
]

NORM_SUMMER = 9.4
NORM_WINTER = 10.3
TANK = 55.0


# =============================================================================
# Fixtures / seed
# =============================================================================


def _fresh_db(tmp_path: Path, name: str = "gsm_season.db") -> str:
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
def registry(repo: GsmRepository) -> GsmRegistryService:
    return GsmRegistryService(repo=repo)


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
    }
    user_id, role = by_user[username]
    return session_cookie(user_id, role, username)


def _seed_vehicle(repo: GsmRepository) -> dict[str, int]:
    driver_id = repo.create_driver(
        full_name="Кулигин Никита Валерьевич",
        license_number="44 21 846315",
    )
    vehicle_id = repo.create_vehicle(
        name="Geely Tugella 848",
        plate_number="О 848 ХР 44",
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        primary_driver_id=driver_id,
    )
    return {"driver_id": driver_id, "vehicle_id": vehicle_id}


def _insert_chain(
    repo: GsmRepository,
    *,
    vehicle_id: int,
    driver_id: int,
    days: list[tuple[date, int, float, str]],
    fuel_start: float = 20.0,
    odometer_start: int = 10_000,
) -> list[int]:
    """Insert waybill chain. days = (date, km, fuel_issued, status)."""
    fuel = fuel_start
    odo = odometer_start
    ids: list[int] = []
    for day, km, issued, status in days:
        burn = burn_for_km(km, NORM_SUMMER)
        fuel_end = round(fuel + issued - burn, 2)
        route_json = json.dumps(
            [{"from": "Завод", "to": f"D {day.isoformat()}", "km": km}],
            ensure_ascii=False,
        )
        wid = repo.upsert_waybill(
            vehicle_id=vehicle_id,
            date=day,
            driver_id=driver_id,
            status=status,
            source="auto",
            odometer_start=odo,
            odometer_end=odo + km,
            fuel_start=fuel,
            fuel_issued=issued,
            fuel_end=fuel_end,
            route_json=route_json,
        )
        ids.append(wid)
        fuel = fuel_end
        odo = odo + km
    return ids


def _tx(day: date, *, service_type: str = "wash", qty_liters: float | None = None) -> Transaction:
    return Transaction(
        card_id=1,
        ts=datetime(day.year, day.month, day.day, 10, 0, 0),
        service_type=service_type,
        qty_liters=qty_liters,
        amount=350.0,
        station_id=10,
        raw_address="station-10",
    )


# =============================================================================
# Core: norm resolution
# =============================================================================


def test_season_mode_defaults_to_summer_without_switches() -> None:
    assert season_mode_for(date(2026, 1, 15), ()) == "summer"
    assert season_mode_for(date(2026, 12, 15), ()) == "summer"
    assert (
        norm_for(
            date(2026, 12, 15),
            norm_summer=NORM_SUMMER,
            norm_winter=NORM_WINTER,
            switches=(),
        )
        == NORM_SUMMER
    )


def test_season_mode_winter_from_switch_date() -> None:
    switches = ((date(2026, 11, 1), "winter"),)
    assert season_mode_for(date(2026, 10, 31), switches) == "summer"
    assert season_mode_for(date(2026, 11, 1), switches) == "winter"
    assert (
        norm_for(
            date(2026, 10, 31),
            norm_summer=NORM_SUMMER,
            norm_winter=NORM_WINTER,
            switches=switches,
        )
        == NORM_SUMMER
    )
    assert (
        norm_for(
            date(2026, 11, 1),
            norm_summer=NORM_SUMMER,
            norm_winter=NORM_WINTER,
            switches=switches,
        )
        == NORM_WINTER
    )


def test_season_mode_switches_back_to_summer() -> None:
    switches = ((date(2026, 11, 1), "winter"), (date(2027, 4, 1), "summer"))
    assert season_mode_for(date(2027, 3, 31), switches) == "winter"
    assert season_mode_for(date(2027, 4, 1), switches) == "summer"
    assert (
        norm_for(
            date(2027, 3, 31),
            norm_summer=NORM_SUMMER,
            norm_winter=NORM_WINTER,
            switches=switches,
        )
        == NORM_WINTER
    )
    assert (
        norm_for(
            date(2027, 4, 1),
            norm_summer=NORM_SUMMER,
            norm_winter=NORM_WINTER,
            switches=switches,
        )
        == NORM_SUMMER
    )


def test_generate_across_switch_boundary_uses_both_norms() -> None:
    """Wash anchors on both sides of the switch burn at their day's norm."""
    summer_day = date(2026, 10, 30)  # Friday, before switch
    winter_day = date(2026, 11, 2)  # Monday, after switch
    km = 120
    routes = (
        LibraryRoute(
            route_id=1,
            addr_a="A",
            addr_b="B",
            km=km,
            frequency=100,
            typical_station_ids=(10,),
        ),
    )
    result = generate(
        transactions=(_tx(summer_day), _tx(winter_day)),
        routes=routes,
        hooks={},
        driver_id=1,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=50.0,
        odometer_start=30_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    by_date = {d.date: d for d in result.days}
    s = by_date[summer_day]
    w = by_date[winter_day]
    daily_km = 2 * km
    summer_burn = s.tank.fuel_start + s.tank.fuel_issued - s.tank.fuel_end
    winter_burn = w.tank.fuel_start + w.tank.fuel_issued - w.tank.fuel_end
    assert summer_burn == pytest.approx(burn_for_km(daily_km, NORM_SUMMER), abs=0.01)
    assert winter_burn == pytest.approx(burn_for_km(daily_km, NORM_WINTER), abs=0.01)


# =============================================================================
# Service: settings wiring
# =============================================================================


def test_service_generate_uses_season_switches_setting(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    station_id = repo.create_station(address="АЗС Тест 1", brand="TATNEFT")
    repo.create_route(
        vehicle_id=ids["vehicle_id"],
        addr_a="Завод",
        addr_b="Объект А",
        km=120,
        frequency=100,
        typical_station_ids=json.dumps([station_id]),
    )
    repo.set_setting(
        "season_switches",
        json.dumps([{"date": "2026-11-01", "mode": "winter"}]),
    )
    batch_id = repo.create_import_batch(
        filename="season.xls",
        uploaded_at="2026-11-05T12:00:00",
        uploaded_by="accountant",
    )
    card_id = repo.create_card(
        card_number="3005454268",
        vehicle_id=ids["vehicle_id"],
        assigned_at="2026-01-01",
    )
    for day in (date(2026, 10, 30), date(2026, 11, 2)):
        repo.insert_transaction(
            card_id=card_id,
            ts=datetime(day.year, day.month, day.day, 10, 0, 0).isoformat(timespec="seconds"),
            service_type="wash",
            qty_liters=None,
            amount=350.0,
            station_id=station_id,
            raw_address="АЗС Тест 1",
            batch_id=batch_id,
        )

    result = service.generate(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2026, 10, 26),
        period_to=date(2026, 11, 6),
        fuel_start=50.0,
        odometer_start=30_000,
    )
    by_date = {w.date: w for w in result.waybills}
    s = by_date["2026-10-30"]
    w = by_date["2026-11-02"]
    assert s.fuel_start + s.fuel_issued - s.fuel_end == pytest.approx(
        burn_for_km(s.km, NORM_SUMMER), abs=0.01
    )
    assert w.fuel_start + w.fuel_issued - w.fuel_end == pytest.approx(
        burn_for_km(w.km, NORM_WINTER), abs=0.01
    )


def test_service_generate_invalid_season_switches_setting(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    repo.set_setting("season_switches", "{not-json")
    with pytest.raises(GsmGenerationError) as exc_info:
        service.generate(
            vehicle_id=ids["vehicle_id"],
            period_from=date(2026, 10, 26),
            period_to=date(2026, 11, 6),
            fuel_start=50.0,
            odometer_start=30_000,
        )
    assert exc_info.value.code == "gsm_settings_invalid"


def test_registry_switch_season_noop_and_monotonic(registry: GsmRegistryService) -> None:
    settings = registry.switch_season(mode="winter", day=date(2026, 11, 1))
    assert settings.season_mode == "winter"
    assert settings.season_switched_at == date(2026, 11, 1)

    # Same mode → no-op, journal not appended.
    again = registry.switch_season(mode="winter", day=date(2026, 12, 1))
    assert again.season_mode == "winter"
    assert again.season_switched_at == date(2026, 11, 1)

    with pytest.raises(GsmRegistryError) as exc_info:
        registry.switch_season(mode="summer", day=date(2026, 10, 15))
    assert exc_info.value.code == "gsm_validation"
    assert "season date must not be before last switch" in str(exc_info.value)

    back = registry.switch_season(mode="summer", day=date(2027, 4, 1))
    assert back.season_mode == "summer"
    assert back.season_switched_at == date(2027, 4, 1)

    with pytest.raises(GsmRegistryError):
        registry.switch_season(mode="autumn", day=date(2027, 5, 1))


# =============================================================================
# HTTP: season endpoint
# =============================================================================


def test_http_season_switch_flow(api_client: CsrfAwareTestClient) -> None:
    cookies = _auth()

    initial = api_client.get(SETTINGS, cookies=cookies)
    assert initial.status_code == 200
    assert initial.json()["season_mode"] == "summer"
    assert initial.json()["season_switched_at"] is None

    switched = api_client.post(
        SEASON,
        json={"mode": "winter", "date": "2026-11-01"},
        cookies=cookies,
    )
    assert switched.status_code == 200, switched.text
    assert switched.json()["season_mode"] == "winter"
    assert switched.json()["season_switched_at"] == "2026-11-01"

    current = api_client.get(SETTINGS, cookies=cookies)
    assert current.json()["season_mode"] == "winter"
    assert current.json()["season_switched_at"] == "2026-11-01"

    noop = api_client.post(
        SEASON,
        json={"mode": "winter", "date": "2026-12-01"},
        cookies=cookies,
    )
    assert noop.status_code == 200
    assert noop.json()["season_switched_at"] == "2026-11-01"

    rejected = api_client.post(
        SEASON,
        json={"mode": "summer", "date": "2026-10-15"},
        cookies=cookies,
    )
    assert rejected.status_code == 422
    detail = rejected.json()["detail"]
    assert detail["code"] == "gsm_validation"
    assert "season date must not be before last switch" in detail["message"]

    back = api_client.post(
        SEASON,
        json={"mode": "summer", "date": "2027-04-01"},
        cookies=cookies,
    )
    assert back.status_code == 200
    assert back.json()["season_mode"] == "summer"
    assert back.json()["season_switched_at"] == "2027-04-01"


def test_http_season_switch_invalid_mode(api_client: CsrfAwareTestClient) -> None:
    resp = api_client.post(
        SEASON,
        json={"mode": "autumn", "date": "2026-11-01"},
        cookies=_auth(),
    )
    assert resp.status_code in (400, 422)


# =============================================================================
# Patch: immutability + cascade re-chain
# =============================================================================


def test_http_patch_confirmed_waybill_locked(api_client: CsrfAwareTestClient) -> None:
    repo = GsmRepository(db_path=str(get_settings().plita_db_path))
    ids = _seed_vehicle(repo)
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "confirmed")],
    )

    resp = api_client.patch(
        f"{WAYBILLS}/{wids[0]}",
        json={"km": 220},
        cookies=_auth(),
    )
    assert resp.status_code == 409
    detail = resp.json()["detail"]
    assert detail["code"] == "gsm_waybill_locked"
    assert "waybill is locked (confirmed/exported)" in detail["message"]

    row = repo.get_waybill_by_id(wids[0])
    assert row is not None
    assert row["status"] == "confirmed"


def test_service_patch_middle_draft_rechains_following(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    days = [
        (date(2025, 4, 7), 190, 40.0, "draft"),
        (date(2025, 4, 8), 180, 0.0, "draft"),
        (date(2025, 4, 9), 200, 0.0, "draft"),
        (date(2025, 4, 10), 160, 30.0, "draft"),
        (date(2025, 4, 11), 200, 0.0, "draft"),
    ]
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=days,
    )

    patched = service.patch_waybill(wids[2], km=220)
    assert patched.rechained_draft_days == 2
    assert patched.km == 220

    after = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 7),
        period_to=date(2025, 4, 11),
    )
    # Chain converges from the edited day onwards; routes/km of followers kept.
    assert after[3].fuel_start == pytest.approx(patched.fuel_end)
    assert after[3].odometer_start == patched.odometer_end
    assert after[3].km == 160
    assert after[4].fuel_start == pytest.approx(after[3].fuel_end)
    assert after[4].odometer_start == after[3].odometer_end
    assert after[4].km == 200
    # Days before the edit are untouched.
    assert after[0].km == 190
    assert after[1].km == 180
    assert after[1].fuel_end == pytest.approx(patched.fuel_start)


def test_http_patch_blocked_by_later_confirmed(api_client: CsrfAwareTestClient) -> None:
    repo = GsmRepository(db_path=str(get_settings().plita_db_path))
    ids = _seed_vehicle(repo)
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[
            (date(2025, 4, 7), 190, 40.0, "draft"),
            (date(2025, 4, 8), 180, 0.0, "confirmed"),
        ],
    )
    before = repo.get_waybill_by_id(wids[0])
    assert before is not None

    resp = api_client.patch(
        f"{WAYBILLS}/{wids[0]}",
        json={"km": 220},
        cookies=_auth(),
    )
    assert resp.status_code == 409
    detail = resp.json()["detail"]
    assert detail["code"] == "gsm_chain_locked"
    assert "cannot edit waybill: later confirmed/exported waybill exists" in detail["message"]

    after = repo.get_waybill_by_id(wids[0])
    assert after is not None
    assert after["status"] == "draft"
    assert json.loads(after["route_json"])[0]["km"] == 190
    assert after["fuel_end"] == pytest.approx(before["fuel_end"])
