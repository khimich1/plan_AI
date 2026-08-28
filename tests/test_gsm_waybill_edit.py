"""Task T11: Правка дня + ручной ПЛ (downstream-пересчёт). TDD.

Acceptance (D11):
- PATCH /gsm/waybills/{id}: route/driver/km → day fixed (source=manual),
  downstream DRAFT days recalculated; confirmed not recalculated
- POST /gsm/waybills: manual constructor with auto fuel fields
- POST /gsm/waybills/{id}/confirm
- After edit km of day N, fuel/odo of N+1… converge
- Manual PL participates in period balance like auto
- REQUIRE_ACCOUNTING; 409/422 as appropriate
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
from core import kp_db_schema
from core.gsm.balance import burn_for_km
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

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

NORM_SUMMER = 9.4
TANK = 55.0


# =============================================================================
# Fixtures / seed
# =============================================================================


def _fresh_db(tmp_path: Path, name: str = "gsm_edit.db") -> str:
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


def _seed_vehicle(repo: GsmRepository) -> dict[str, int]:
    driver_id = repo.create_driver(
        full_name="Кулигин Никита Валерьевич",
        license_number="44 21 846315",
    )
    driver2_id = repo.create_driver(
        full_name="Скрябин Иван",
        license_number="44 21 111111",
    )
    vehicle_id = repo.create_vehicle(
        name="Geely Tugella 848",
        plate_number="О 848 ХР 44",
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=10.3,
        primary_driver_id=driver_id,
    )
    card_id = repo.create_card(
        card_number="3005454268",
        vehicle_id=vehicle_id,
        assigned_at="2025-01-01",
    )
    station_id = repo.create_station(address="АЗС Тест 1", brand="TATNEFT")
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
    repo.set_setting("winter_start", "11-01")
    repo.set_setting("hook_threshold_km", "13")
    return {
        "driver_id": driver_id,
        "driver2_id": driver2_id,
        "vehicle_id": vehicle_id,
        "card_id": card_id,
        "station_id": station_id,
    }


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


def _assert_chain_converges(waybills: list, *, norm: float = NORM_SUMMER) -> None:
    ordered = sorted(waybills, key=lambda w: w.date if hasattr(w, "date") else w["date"])
    for i, wb in enumerate(ordered):
        date_s = wb.date if hasattr(wb, "date") else wb["date"]
        fuel_start = wb.fuel_start if hasattr(wb, "fuel_start") else wb["fuel_start"]
        fuel_issued = wb.fuel_issued if hasattr(wb, "fuel_issued") else wb["fuel_issued"]
        fuel_end = wb.fuel_end if hasattr(wb, "fuel_end") else wb["fuel_end"]
        km = wb.km if hasattr(wb, "km") else wb.get("km")
        if km is None:
            route = wb.route if hasattr(wb, "route") else json.loads(wb.get("route_json") or "[]")
            if hasattr(route[0], "km"):
                km = sum(leg.km for leg in route)
            else:
                km = sum(int(leg.get("km") or 0) for leg in route)
        odo_s = wb.odometer_start if hasattr(wb, "odometer_start") else wb["odometer_start"]
        odo_e = wb.odometer_end if hasattr(wb, "odometer_end") else wb["odometer_end"]

        burn = burn_for_km(int(km), norm)
        assert fuel_end == pytest.approx(round(float(fuel_start) + float(fuel_issued) - burn, 2)), date_s
        assert int(odo_e) == int(odo_s) + int(km), date_s
        if i > 0:
            prev = ordered[i - 1]
            prev_end = prev.fuel_end if hasattr(prev, "fuel_end") else prev["fuel_end"]
            prev_odo = prev.odometer_end if hasattr(prev, "odometer_end") else prev["odometer_end"]
            assert float(fuel_start) == pytest.approx(float(prev_end)), f"fuel link {date_s}"
            assert int(odo_s) == int(prev_odo), f"odo link {date_s}"


def _seed_into_api_db(api_client: CsrfAwareTestClient) -> dict[str, int]:
    from app.core.settings import get_settings as _gs

    return _seed_vehicle(GsmRepository(db_path=str(_gs().plita_db_path)))


# =============================================================================
# Service-level
# =============================================================================


def test_service_patch_km_recalculates_downstream_drafts(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    # Mon Tue Wed — all draft
    days = [
        (date(2025, 4, 7), 190, 40.0, "draft"),
        (date(2025, 4, 8), 180, 0.0, "draft"),
        (date(2025, 4, 9), 200, 0.0, "draft"),
    ]
    wids = _insert_chain(repo, vehicle_id=ids["vehicle_id"], driver_id=ids["driver_id"], days=days)
    day_n_id = wids[0]

    before = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 7),
        period_to=date(2025, 4, 9),
    )
    assert len(before) == 3
    old_n1_fuel_start = before[1].fuel_start
    old_n1_odo_start = before[1].odometer_start

    patched = service.patch_waybill(
        day_n_id,
        km=220,
    )
    assert patched.source == "manual"
    assert patched.km == 220
    assert patched.rechained_draft_days == 2
    assert patched.odometer_end == patched.odometer_start + 220
    burn = burn_for_km(220, NORM_SUMMER)
    assert patched.fuel_end == pytest.approx(
        round(float(patched.fuel_start) + float(patched.fuel_issued) - burn, 2)
    )

    after = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 7),
        period_to=date(2025, 4, 9),
    )
    _assert_chain_converges(after)
    assert after[1].fuel_start != old_n1_fuel_start or after[1].odometer_start != old_n1_odo_start
    assert after[1].fuel_start == pytest.approx(patched.fuel_end)
    assert after[1].odometer_start == patched.odometer_end
    assert after[1].km == 180  # route/km of downstream kept
    assert after[2].fuel_start == pytest.approx(after[1].fuel_end)


def test_service_patch_rejected_when_confirmed_downstream(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    """A later confirmed/exported day locks the whole chain above it (D11b)."""
    ids = _seed_vehicle(repo)
    days = [
        (date(2025, 4, 7), 190, 40.0, "draft"),
        (date(2025, 4, 8), 180, 0.0, "confirmed"),
        (date(2025, 4, 9), 200, 0.0, "draft"),
    ]
    wids = _insert_chain(repo, vehicle_id=ids["vehicle_id"], driver_id=ids["driver_id"], days=days)

    before = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 7),
        period_to=date(2025, 4, 9),
    )

    with pytest.raises(GsmGenerationError) as exc_info:
        service.patch_waybill(wids[0], km=220)
    assert exc_info.value.code == "gsm_chain_locked"
    assert "later confirmed/exported waybill exists" in str(exc_info.value)

    # Whole edit rejected: nothing persisted.
    after = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 7),
        period_to=date(2025, 4, 9),
    )
    for prev, cur in zip(before, after, strict=True):
        assert cur.km == prev.km
        assert cur.fuel_start == pytest.approx(prev.fuel_start)
        assert cur.fuel_end == pytest.approx(prev.fuel_end)
        assert cur.odometer_start == prev.odometer_start
        assert cur.odometer_end == prev.odometer_end
        assert cur.status == prev.status


def test_service_patch_driver_and_route(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "draft")],
    )
    patched = service.patch_waybill(
        wids[0],
        driver_id=ids["driver2_id"],
        route=[
            {"from": "Завод", "to": "Новый объект", "km": 175, "station_id": ids["station_id"]},
        ],
    )
    assert patched.driver_id == ids["driver2_id"]
    assert patched.source == "manual"
    assert patched.km == 175
    assert patched.route[0].to_addr == "Новый объект"
    assert patched.route[0].station_id == ids["station_id"]


def test_service_patch_corridor_violation_422(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    # Low fuel, huge km → negative fuel_end
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 100, 0.0, "draft")],
        fuel_start=5.0,
        odometer_start=10_000,
    )
    with pytest.raises(GsmGenerationError) as exc_info:
        service.patch_waybill(wids[0], km=250)
    assert exc_info.value.code in {"gsm_unsolvable", "gsm_balance_violation"}


def test_service_create_manual_auto_fuel_and_downstream(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    # Existing day after the gap we will fill manually
    _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[
            (date(2025, 4, 7), 190, 40.0, "draft"),
            (date(2025, 4, 9), 180, 0.0, "draft"),
        ],
        fuel_start=20.0,
        odometer_start=10_000,
    )

    created = service.create_waybill(
        vehicle_id=ids["vehicle_id"],
        day=date(2025, 4, 8),
        driver_id=ids["driver_id"],
        route=[{"from": "Завод", "to": "Ручной", "km": 160}],
        fuel_issued=0.0,
    )
    assert created.source == "manual"
    assert created.status == "draft"
    assert created.date == "2025-04-08"
    assert created.km == 160
    # Auto fuel: start from previous day's fuel_end
    prev = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 7),
        period_to=date(2025, 4, 7),
    )[0]
    assert created.fuel_start == pytest.approx(prev.fuel_end)
    assert created.odometer_start == prev.odometer_end
    burn = burn_for_km(160, NORM_SUMMER)
    assert created.fuel_end == pytest.approx(
        round(float(created.fuel_start) + 0.0 - burn, 2)
    )

    all_days = service.list_waybills(
        vehicle_id=ids["vehicle_id"],
        period_from=date(2025, 4, 7),
        period_to=date(2025, 4, 9),
    )
    assert len(all_days) == 3
    _assert_chain_converges(all_days)
    # Manual participates: Apr 9 starts from manual's end
    assert all_days[2].fuel_start == pytest.approx(created.fuel_end)
    assert all_days[2].odometer_start == created.odometer_end


def test_service_create_manual_duplicate_date_conflict(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "draft")],
    )
    with pytest.raises(GsmGenerationError) as exc_info:
        service.create_waybill(
            vehicle_id=ids["vehicle_id"],
            day=date(2025, 4, 7),
            driver_id=ids["driver_id"],
            route=[{"from": "A", "to": "B", "km": 150}],
        )
    assert exc_info.value.code == "gsm_waybill_conflict"


def test_service_confirm_waybill(
    service: GsmGenerationService, repo: GsmRepository
) -> None:
    ids = _seed_vehicle(repo)
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "draft")],
    )
    confirmed = service.confirm_waybill(wids[0])
    assert confirmed.status == "confirmed"
    assert confirmed.id == wids[0]

    row = repo.get_waybill_by_id(wids[0])
    assert row is not None
    assert row["status"] == "confirmed"


# =============================================================================
# HTTP-level
# =============================================================================


def test_http_patch_recalculates_downstream(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth()
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[
            (date(2025, 4, 7), 190, 40.0, "draft"),
            (date(2025, 4, 8), 180, 0.0, "draft"),
            (date(2025, 4, 9), 200, 0.0, "draft"),
        ],
    )

    resp = api_client.patch(
        f"{WAYBILLS}/{wids[0]}",
        json={"km": 220},
        cookies=cookies,
    )
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["source"] == "manual"
    assert body["km"] == 220

    listed = api_client.get(
        WAYBILLS,
        params={"vehicle_id": ids["vehicle_id"], "from": "2025-04-07", "to": "2025-04-09"},
        cookies=cookies,
    )
    assert listed.status_code == 200
    days = listed.json()
    assert days[1]["fuel_start"] == pytest.approx(body["fuel_end"])
    assert days[1]["odometer_start"] == body["odometer_end"]
    assert days[2]["fuel_start"] == pytest.approx(days[1]["fuel_end"])


def test_http_create_manual_and_confirm(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth()
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "draft")],
        fuel_start=20.0,
        odometer_start=10_000,
    )

    created = api_client.post(
        WAYBILLS,
        json={
            "vehicle_id": ids["vehicle_id"],
            "date": "2025-04-08",
            "driver_id": ids["driver_id"],
            "route": [{"from": "Завод", "to": "Ручной", "km": 160}],
            "fuel_issued": 0.0,
        },
        cookies=cookies,
    )
    assert created.status_code == 201, created.text
    body = created.json()
    assert body["source"] == "manual"
    assert body["fuel_start"] is not None
    assert body["fuel_end"] is not None
    assert body["odometer_start"] is not None
    assert body["odometer_end"] == body["odometer_start"] + 160

    conf = api_client.post(f"{WAYBILLS}/{body['id']}/confirm", cookies=cookies)
    assert conf.status_code == 200, conf.text
    assert conf.json()["status"] == "confirmed"


def test_http_create_duplicate_409(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth()
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "draft")],
    )
    resp = api_client.post(
        WAYBILLS,
        json={
            "vehicle_id": ids["vehicle_id"],
            "date": "2025-04-07",
            "driver_id": ids["driver_id"],
            "route": [{"from": "A", "to": "B", "km": 150}],
        },
        cookies=cookies,
    )
    assert resp.status_code == 409
    assert resp.json()["detail"]["code"] == "gsm_waybill_conflict"


def test_http_patch_balance_violation_422(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth()
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 100, 0.0, "draft")],
        fuel_start=5.0,
    )
    resp = api_client.patch(
        f"{WAYBILLS}/{wids[0]}",
        json={"km": 250},
        cookies=cookies,
    )
    assert resp.status_code == 422


def test_http_manager_forbidden_edit(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth("manager_a")
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "draft")],
    )

    assert (
        api_client.patch(
            f"{WAYBILLS}/{wids[0]}",
            json={"km": 200},
            cookies=cookies,
        ).status_code
        == 403
    )
    assert (
        api_client.post(
            WAYBILLS,
            json={
                "vehicle_id": ids["vehicle_id"],
                "date": "2025-04-08",
                "driver_id": ids["driver_id"],
                "route": [{"from": "A", "to": "B", "km": 150}],
            },
            cookies=cookies,
        ).status_code
        == 403
    )
    assert (
        api_client.post(f"{WAYBILLS}/{wids[0]}/confirm", cookies=cookies).status_code
        == 403
    )


def test_http_accountant_allowed_patch(api_client: CsrfAwareTestClient) -> None:
    ids = _seed_into_api_db(api_client)
    cookies = _auth()
    from app.core.settings import get_settings as _gs

    repo = GsmRepository(db_path=str(_gs().plita_db_path))
    wids = _insert_chain(
        repo,
        vehicle_id=ids["vehicle_id"],
        driver_id=ids["driver_id"],
        days=[(date(2025, 4, 7), 190, 40.0, "draft")],
    )
    resp = api_client.patch(
        f"{WAYBILLS}/{wids[0]}",
        json={"driver_id": ids["driver_id"], "km": 195},
        cookies=cookies,
    )
    assert resp.status_code == 200, resp.text
