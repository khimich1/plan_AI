"""Unit tests for core.gsm.generator (pure domain, no app I/O).

TDD / Task 9: these tests are written FIRST and must fail until
``core/gsm/generator.py`` exists.

Worker note (balance rounding)
------------------------------
Prefer ``round(fuel_end, 2)`` inside ``core.gsm.balance.apply_day`` so the
period identity ``Σ issued − Σ burn = Δfuel`` holds at 2 decimal places
without float drift across long chains. Generator tests use
``pytest.approx(..., abs=0.01)`` either way; rounding in balance is the
preferred fix (T8 deferred this to the generator path).

Decisions pinned here
---------------------
D4  Anchors = fuel OR wash → waybill that day; burn-in only on weekdays
    without RF holidays (holidays injected); weekend/holiday tx → still a
    day + warning ``weekend_anchor``.
D5  Corridor ``[0…tank]`` every day; before issue ``Q`` require
    ``fuel_start ≤ tank − Q`` (retro burn-in); if impossible → draft day
    with ``manual_intervention`` in ``problematic_days`` (period continues;
    ``unsolvable`` stays None for balance/headroom/corridor).
D6  Season from manual ``season_switches`` journal (not calendar auto):
    mode(day) = mode of the latest switch ≤ day (default summer) →
    ``norm_winter`` iff mode is winter, else ``norm_summer``.
D12 New station → min ``крюк_км``; hook > ``hook_threshold_km`` (default 13)
    → warning ``hook_above_threshold``.

Public surface pinned by this file
----------------------------------
``core.gsm.generator``
    ``LibraryRoute`` — frozen library entry:
        ``route_id, addr_a, addr_b, km, frequency, typical_station_ids``,
        ``vehicle_id`` (default 0), optional ``point_a`` / ``point_b``
        (``GeoPoint``) for direction sort.
    ``UnsolvableInfo`` — kept for compatibility (no longer filled for
        balance/headroom/corridor).
    ``ProblematicDay`` — local unsolvable anchor: ``date``, ``reason``,
        ``detail``, ``fuel_before``, ``fuel_to_issue``, ``tank_volume``.
    ``GenerateResult`` — frozen result:
        ``days: tuple[WaybillDay, ...]``
        ``unsolvable: UnsolvableInfo | None``
        ``warnings: tuple[str, ...]``  (period-level, deduped codes)
        ``problematic_days: tuple[ProblematicDay, ...]``
    ``generate(*, transactions, routes, hooks, driver_id,
               tank_volume_liters, norm_summer, norm_winter, season_switches,
               fuel_start, odometer_start, hook_threshold_km=13.0,
               holidays=..., extra_workdays=..., seed=0,
               max_daily_km=700, station_coords=None,
               own_vehicle_id=0) -> GenerateResult``
        ``hooks`` maps ``(route_id, station_id) -> крюк_км`` (precomputed;
        no OSRM/I/O in core). Burn-in candidates are routes already within
        ``max_daily_km`` (``2×km ≤ cap``), ordered by frequency desc (ties:
        lower route_id, then seed). When a burn already lands in headroom,
        pick min sufficient daily km (then frequency desc, ``route_id``).
        Anchor/burn-in candidates with ``2×km > max_daily_km`` are dropped.
"""

from __future__ import annotations

import ast
import dataclasses
import inspect
from datetime import date, datetime
from pathlib import Path

import pytest

from core.gsm.balance import burn_for_km
from core.gsm.generator import (
    GenerateResult,
    LibraryRoute,
    ProblematicDay,
    UnsolvableInfo,
    _city_key,
    _find_home_twin,
    _is_home_base,
    _orient_home_round_trip,
    generate,
)
from core.gsm.geo import GeoPoint
from core.gsm.models import RouteRef, Transaction, WaybillDay

REPO_ROOT = Path(__file__).resolve().parents[1]
GSM_GENERATOR_PATH = REPO_ROOT / "core" / "gsm" / "generator.py"

# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------

TANK = 55.0
NORM_SUMMER = 9.4
NORM_WINTER = 10.3
DRIVER_ID = 1
CARD_ID = 1


def _tx(
    day: date,
    *,
    service_type: str = "fuel",
    qty_liters: float | None = 40.0,
    station_id: int = 10,
    hour: int = 10,
) -> Transaction:
    return Transaction(
        card_id=CARD_ID,
        ts=datetime(day.year, day.month, day.day, hour, 0, 0),
        service_type=service_type,
        qty_liters=qty_liters,
        amount=100.0 if service_type == "fuel" else 350.0,
        station_id=station_id,
        raw_address=f"station-{station_id}",
    )


def _route(
    route_id: int,
    *,
    km: int = 190,
    frequency: int = 10,
    typical_station_ids: tuple[int, ...] = (),
    addr_a: str = "A",
    addr_b: str = "B",
    point_a: GeoPoint | None = None,
    point_b: GeoPoint | None = None,
    vehicle_id: int = 0,
) -> LibraryRoute:
    return LibraryRoute(
        route_id=route_id,
        addr_a=addr_a,
        addr_b=addr_b,
        km=km,
        frequency=frequency,
        typical_station_ids=typical_station_ids,
        vehicle_id=vehicle_id,
        point_a=point_a,
        point_b=point_b,
    )


def _burn_library() -> tuple[LibraryRoute, ...]:
    """Routes in the burn-in band [150, 250] with distinct frequencies."""
    return (
        _route(101, km=200, frequency=50, typical_station_ids=()),
        _route(102, km=180, frequency=40, typical_station_ids=()),
        _route(103, km=220, frequency=30, typical_station_ids=()),
        _route(104, km=150, frequency=20, typical_station_ids=()),
        _route(105, km=250, frequency=10, typical_station_ids=()),
    )


def _day_by_date(result: GenerateResult, day: date) -> WaybillDay:
    matches = [d for d in result.days if d.date == day]
    assert len(matches) == 1, f"expected exactly one waybill on {day}, got {len(matches)}"
    return matches[0]


def _norm_for(day: date, switches: tuple[tuple[date, str], ...]) -> float:
    mode = "summer"
    for switch_date, switch_mode in switches:
        if switch_date > day:
            break
        mode = switch_mode
    return NORM_WINTER if mode == "winter" else NORM_SUMMER


def _period_burn(result: GenerateResult, switches: tuple[tuple[date, str], ...]) -> float:
    total = 0.0
    for day in result.days:
        total += burn_for_km(day.tank.km, _norm_for(day.date, switches))
    return total


def _period_issued(result: GenerateResult) -> float:
    return sum(d.tank.fuel_issued for d in result.days)


# ---------------------------------------------------------------------------
# 1. Synthetic period: fuel/wash anchors + station on selected route
# ---------------------------------------------------------------------------


def test_generate_anchors_fuel_and_wash_with_station_route() -> None:
    """Each fuel/wash tx → a waybill day; selected route lists that station."""
    fuel_day = date(2025, 4, 7)  # Monday
    wash_day = date(2025, 4, 9)  # Wednesday

    station_fuel = 10
    station_wash = 20
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(station_fuel,)),
        _route(2, km=180, frequency=90, typical_station_ids=(station_wash,)),
        *_burn_library(),
    )
    txs = (
        _tx(fuel_day, service_type="fuel", qty_liters=35.0, station_id=station_fuel),
        _tx(wash_day, service_type="wash", qty_liters=None, station_id=station_wash),
    )

    result = generate(
        transactions=txs,
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=36.0,
        odometer_start=100_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )

    assert isinstance(result, GenerateResult)
    assert result.unsolvable is None
    assert {d.date for d in result.days} >= {fuel_day, wash_day}

    fuel_wb = _day_by_date(result, fuel_day)
    wash_wb = _day_by_date(result, wash_day)
    assert fuel_wb.tank.fuel_issued == pytest.approx(35.0)
    assert wash_wb.tank.fuel_issued == pytest.approx(0.0)
    assert fuel_wb.route.route_id == 1
    assert wash_wb.route.route_id == 2
    assert fuel_wb.source == "auto"
    assert wash_wb.source == "auto"
    assert fuel_wb.driver_id == DRIVER_ID


def test_generate_multiple_txs_same_day_single_waybill() -> None:
    """Several txs on one calendar day collapse to one WaybillDay."""
    day = date(2025, 4, 7)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10, 11)),
        *_burn_library(),
    )
    txs = (
        _tx(day, qty_liters=20.0, station_id=10, hour=9),
        _tx(day, qty_liters=15.0, station_id=11, hour=15),
    )
    result = generate(
        transactions=txs,
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=10.0,
        odometer_start=50_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert sum(1 for d in result.days if d.date == day) == 1
    assert _day_by_date(result, day).tank.fuel_issued == pytest.approx(35.0)


# ---------------------------------------------------------------------------
# 2. Burn-in weekdays only; weekend tx → weekend_anchor
# ---------------------------------------------------------------------------


def test_burn_in_skips_weekends_and_injected_holidays() -> None:
    """Burn-in fills only working days; holiday mid-week is never a burn day."""
    # Fri 4 Apr fuel → Mon 14 Apr fuel (room for burn-in).
    first = date(2025, 4, 4)  # Friday
    second = date(2025, 4, 14)  # Monday
    holiday = date(2025, 4, 9)  # Wednesday injected as RF holiday

    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        *_burn_library(),
    )
    txs = (
        _tx(first, qty_liters=45.0, station_id=10),
        _tx(second, qty_liters=40.0, station_id=10),
    )
    result = generate(
        transactions=txs,
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=22.0,
        odometer_start=80_000,
        holidays=frozenset({holiday}),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None

    burn_dates = {
        d.date
        for d in result.days
        if d.date not in (first, second) and abs(d.tank.fuel_issued) < 1e-9
    }
    assert burn_dates, "expected burn-in days between fuels"
    for d in burn_dates:
        assert d.weekday() < 5, f"burn-in on weekend {d}"
        assert d != holiday, f"burn-in on holiday {d}"
    # Injected holiday must not appear as a generated day at all.
    assert holiday not in {d.date for d in result.days}


def test_weekend_transaction_is_anchor_with_weekend_anchor_warning() -> None:
    """Weekend tx still creates a day and emits weekend_anchor (D4)."""
    sunday = date(2025, 4, 6)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        *_burn_library(),
    )
    result = generate(
        transactions=(_tx(sunday, qty_liters=30.0, station_id=10),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=10_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    wb = _day_by_date(result, sunday)
    assert "weekend_anchor" in wb.warnings
    assert "weekend_anchor" in result.warnings


def test_holiday_transaction_gets_weekend_anchor_warning() -> None:
    """RF holiday (injected) counts as non-working → weekend_anchor."""
    wed = date(2025, 4, 9)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        *_burn_library(),
    )
    result = generate(
        transactions=(_tx(wed, qty_liters=25.0, station_id=10),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=18.0,
        odometer_start=20_000,
        holidays=frozenset({wed}),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    wb = _day_by_date(result, wed)
    assert "weekend_anchor" in wb.warnings


# ---------------------------------------------------------------------------
# 3. Period fuel identity + daily corridor
# ---------------------------------------------------------------------------


def test_period_fuel_identity_and_daily_corridor() -> None:
    """Σ issued − Σ burn = fuel_end_last − fuel_start_first; corridor every day."""
    fuel_day = date(2025, 4, 7)
    wash_day = date(2025, 4, 11)
    switches = ((date(2025, 11, 1), "winter"),)
    fuel_start = 32.0

    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        _route(2, km=180, frequency=80, typical_station_ids=(20,)),
        *_burn_library(),
    )
    txs = (
        _tx(fuel_day, qty_liters=40.0, station_id=10),
        _tx(wash_day, service_type="wash", qty_liters=None, station_id=20),
    )
    result = generate(
        transactions=txs,
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=switches,
        fuel_start=fuel_start,
        odometer_start=70_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert len(result.days) >= 2

    for day in result.days:
        assert 0.0 <= day.tank.fuel_end <= TANK
        assert 0.0 <= day.tank.fuel_start <= TANK
        assert isinstance(day.route, RouteRef)
        assert day.tank.km == 2 * day.route.km
        assert day.tank.km == sum(leg.km for leg in day.legs)
        assert day.tank.odometer_end == day.tank.odometer_start + day.tank.km

    # Continuity
    ordered = sorted(result.days, key=lambda d: d.date)
    for prev, cur in zip(ordered[:-1], ordered[1:], strict=True):
        assert cur.tank.fuel_start == pytest.approx(prev.tank.fuel_end, abs=0.01)
        assert cur.tank.odometer_start == prev.tank.odometer_end

    issued = _period_issued(result)
    burned = _period_burn(result, switches)
    delta = ordered[-1].tank.fuel_end - fuel_start
    assert issued - burned == pytest.approx(delta, abs=0.01)


# ---------------------------------------------------------------------------
# 4. Headroom before fuel Q (retro burn) / unsolvable
# ---------------------------------------------------------------------------


def test_retro_burn_creates_headroom_before_large_fuel() -> None:
    """Before issue Q, fuel_start ≤ tank − Q (D5); retro burn-in makes it so."""
    first = date(2025, 4, 7)  # Monday
    second = date(2025, 4, 18)  # Friday — enough weekdays between
    q_second = 45.0
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        *_burn_library(),
    )
    txs = (
        _tx(first, qty_liters=45.0, station_id=10),
        _tx(second, qty_liters=q_second, station_id=10),
    )
    result = generate(
        transactions=txs,
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=90_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    second_wb = _day_by_date(result, second)
    headroom_needed = TANK - q_second
    assert second_wb.tank.fuel_start <= headroom_needed + 1e-9
    assert second_wb.tank.fuel_start + q_second - burn_for_km(
        second_wb.tank.km, NORM_SUMMER
    ) <= TANK + 1e-9
    # At least one burn-in day between anchors.
    between = [d for d in result.days if first < d.date < second]
    assert between, "expected retro/gap burn-in days between fuels"


def test_unsolvable_when_no_weekdays_for_headroom() -> None:
    """Fri→Mon overflow with no free weekdays → period kept, Monday is manual."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    q = 40.0
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        _route(2, km=150, frequency=40, typical_station_ids=()),
    )
    # Start high enough that Friday fill leaves tank too full for Monday's Q,
    # and Sat/Sun cannot be used for burn-in. Library has no 2×km ≥ km_needed.
    txs = (
        _tx(friday, qty_liters=40.0, station_id=10),
        _tx(monday, qty_liters=q, station_id=10),
    )
    result = generate(
        transactions=txs,
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert {d.date for d in result.days} >= {friday, monday}
    monday_wb = _day_by_date(result, monday)
    assert monday_wb.source == "auto"
    assert "manual_intervention" in monday_wb.warnings
    problems = [p for p in result.problematic_days if p.date == monday]
    assert len(problems) == 1
    problem = problems[0]
    assert problem.reason == "manual_intervention"
    assert problem.fuel_to_issue == pytest.approx(q)
    assert problem.tank_volume == pytest.approx(TANK)
    assert problem.detail


# ---------------------------------------------------------------------------
# 5. Season via manual season_switches (D6)
# ---------------------------------------------------------------------------


def test_season_switches_at_switch_date() -> None:
    """Same km: date before winter switch uses summer norm; on/after → winter."""
    summer_day = date(2025, 10, 31)  # Friday
    winter_day = date(2025, 11, 3)  # Monday
    switches = ((date(2025, 11, 1), "winter"),)
    km = 120
    routes = (
        _route(1, km=km, frequency=100, typical_station_ids=(10,)),
        _route(2, km=km, frequency=90, typical_station_ids=(20,)),
        *_burn_library(),
    )
    # Washes → issued 0 so burn is visible as fuel_start − fuel_end.
    txs = (
        _tx(summer_day, service_type="wash", qty_liters=None, station_id=10),
        _tx(winter_day, service_type="wash", qty_liters=None, station_id=20),
    )
    result = generate(
        transactions=txs,
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=switches,
        fuel_start=50.0,
        odometer_start=30_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    s = _day_by_date(result, summer_day)
    w = _day_by_date(result, winter_day)
    daily_km = 2 * km
    # Library km is one leg; tank burns the round-trip.
    assert s.route.km == km
    assert w.route.km == km
    assert s.tank.km == daily_km
    assert w.tank.km == daily_km
    summer_burn = s.tank.fuel_start + s.tank.fuel_issued - s.tank.fuel_end
    winter_burn = w.tank.fuel_start + w.tank.fuel_issued - w.tank.fuel_end
    assert summer_burn == pytest.approx(burn_for_km(daily_km, NORM_SUMMER), abs=0.01)
    assert winter_burn == pytest.approx(burn_for_km(daily_km, NORM_WINTER), abs=0.01)
    assert winter_burn > summer_burn


# ---------------------------------------------------------------------------
# 6. Min hook > threshold → hook_above_threshold (D12)
# ---------------------------------------------------------------------------


def test_hook_above_threshold_warning_for_new_station() -> None:
    """Station not typical → min крюк; if > threshold → hook_above_threshold."""
    day = date(2025, 4, 8)  # Tuesday
    new_station = 99
    routes = (
        _route(1, km=190, frequency=50, typical_station_ids=(10,)),
        _route(2, km=200, frequency=40, typical_station_ids=(11,)),
        _route(3, km=180, frequency=30, typical_station_ids=()),
        *_burn_library(),
    )
    hooks = {
        (1, new_station): 20.0,
        (2, new_station): 15.5,  # min among hooks, still > 13
        (3, new_station): 18.0,
    }
    result = generate(
        transactions=(_tx(day, qty_liters=30.0, station_id=new_station),),
        routes=routes,
        hooks=hooks,
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=15.0,
        odometer_start=40_000,
        hook_threshold_km=13.0,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    wb = _day_by_date(result, day)
    assert wb.route.route_id == 2  # min hook 15.5
    assert "hook_above_threshold" in wb.warnings
    assert "hook_above_threshold" in result.warnings


def test_hook_within_threshold_no_warning() -> None:
    day = date(2025, 4, 8)
    new_station = 99
    routes = (
        _route(1, km=190, frequency=50, typical_station_ids=()),
        _route(2, km=200, frequency=40, typical_station_ids=()),
        *_burn_library(),
    )
    hooks = {
        (1, new_station): 12.0,
        (2, new_station): 4.5,  # min, under threshold
    }
    result = generate(
        transactions=(_tx(day, qty_liters=30.0, station_id=new_station),),
        routes=routes,
        hooks=hooks,
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=15.0,
        odometer_start=40_000,
        hook_threshold_km=13.0,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    wb = _day_by_date(result, day)
    assert wb.route.route_id == 2
    assert "hook_above_threshold" not in wb.warnings


# ---------------------------------------------------------------------------
# 7. Determinism
# ---------------------------------------------------------------------------


def test_generate_is_deterministic() -> None:
    first = date(2025, 4, 7)
    second = date(2025, 4, 16)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        *_burn_library(),
    )
    kwargs = dict(
        transactions=(
            _tx(first, qty_liters=42.0, station_id=10),
            _tx(second, qty_liters=38.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=25.0,
        odometer_start=55_000,
        holidays=frozenset({date(2025, 4, 10)}),
        extra_workdays=frozenset(),
        seed=7,
    )
    a = generate(**kwargs)
    b = generate(**kwargs)
    assert a == b
    assert a.unsolvable is None


# ---------------------------------------------------------------------------
# 8. AST guard: no app.* imports
# ---------------------------------------------------------------------------


def test_gsm_generator_module_has_no_app_imports() -> None:
    assert GSM_GENERATOR_PATH.is_file(), f"missing module file: {GSM_GENERATOR_PATH}"
    tree = ast.parse(GSM_GENERATOR_PATH.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert not alias.name.startswith("app"), alias.name
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            assert not module.startswith("app"), module


# ---------------------------------------------------------------------------
# 9. Result / input DTO shape
# ---------------------------------------------------------------------------


def test_generate_result_api_fields() -> None:
    """Pin GenerateResult API: days, unsolvable, warnings, problematic_days."""
    for cls in (GenerateResult, LibraryRoute, UnsolvableInfo, ProblematicDay):
        assert dataclasses.is_dataclass(cls), f"{cls.__name__} must be a dataclass"
        params = getattr(cls, "__dataclass_params__", None)
        assert params is not None and params.frozen, f"{cls.__name__} must be frozen=True"
        assert hasattr(cls, "__slots__"), f"{cls.__name__} must use slots=True"

    fields = {f.name for f in dataclasses.fields(GenerateResult)}
    assert fields == {"days", "unsolvable", "warnings", "problematic_days"}

    route_fields = {f.name for f in dataclasses.fields(LibraryRoute)}
    assert "vehicle_id" in route_fields
    vehicle_id_field = next(f for f in dataclasses.fields(LibraryRoute) if f.name == "vehicle_id")
    assert vehicle_id_field.default == 0

    problem_fields = {f.name for f in dataclasses.fields(ProblematicDay)}
    for required in (
        "date",
        "reason",
        "detail",
        "fuel_before",
        "fuel_to_issue",
        "tank_volume",
    ):
        assert required in problem_fields, required

    info_fields = {f.name for f in dataclasses.fields(UnsolvableInfo)}
    for required in (
        "reason",
        "at_date",
        "fuel_before",
        "fuel_to_issue",
        "tank_volume",
        "free_weekdays",
        "detail",
    ):
        assert required in info_fields, required


# ---------------------------------------------------------------------------
# 10. Round-trip: day = two legs, daily km = 2× library km
# ---------------------------------------------------------------------------


def _assert_round_trip_day(wb: WaybillDay, *, library_km: int, norm: float) -> None:
    assert len(wb.legs) == 2
    out, back = wb.legs
    assert out.route_id == wb.route.route_id
    assert back.route_id == wb.route.route_id
    assert out.addr_a == wb.route.addr_a
    assert out.addr_b == wb.route.addr_b
    assert back.addr_a == wb.route.addr_b
    assert back.addr_b == wb.route.addr_a
    assert out.km == library_km
    assert back.km == library_km
    assert wb.route.km == library_km
    assert wb.tank.km == 2 * library_km
    assert wb.tank.km == out.km + back.km
    burned = wb.tank.fuel_start + wb.tank.fuel_issued - wb.tank.fuel_end
    assert burned == pytest.approx(burn_for_km(2 * library_km, norm), abs=0.01)


def test_anchor_day_is_round_trip_two_legs() -> None:
    """Anchor burns 2× library km and stores A→B + B→A legs."""
    day = date(2025, 4, 7)
    library_km = 190
    routes = (
        _route(
            1,
            km=library_km,
            frequency=100,
            typical_station_ids=(10,),
            addr_a="Завод",
            addr_b="Объект",
        ),
    )
    result = generate(
        transactions=(_tx(day, qty_liters=40.0, station_id=10),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=10_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    wb = _day_by_date(result, day)
    _assert_round_trip_day(wb, library_km=library_km, norm=NORM_SUMMER)


def test_burn_in_day_is_round_trip_two_legs() -> None:
    """Retro burn-in days are also round-trips (2× km / two legs)."""
    first = date(2025, 4, 7)
    second = date(2025, 4, 18)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,), addr_a="A", addr_b="B"),
        *_burn_library(),
    )
    result = generate(
        transactions=(
            _tx(first, qty_liters=45.0, station_id=10),
            _tx(second, qty_liters=45.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=90_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    between = [d for d in result.days if first < d.date < second]
    assert between, "expected burn-in days between fuels"
    for wb in between:
        _assert_round_trip_day(wb, library_km=wb.route.km, norm=NORM_SUMMER)
    _assert_round_trip_day(_day_by_date(result, first), library_km=190, norm=NORM_SUMMER)
    _assert_round_trip_day(_day_by_date(result, second), library_km=190, norm=NORM_SUMMER)


# ---------------------------------------------------------------------------
# 11. Lookahead on the anchor (Task 5): elongate only when needed
# ---------------------------------------------------------------------------

# Fri→Mon, no free weekdays. TANK=55, NORM=9.4, Q=40, start=20:
#   headroom_needed = 55 − 40 = 15
#   fuel_after      = 20 + 40 = 60
#   burn_needed     = 60 − 15 = 45
#   km_needed       = 45 / 9.4 × 100 ≈ 478.72
# 190 → daily 380, burn ≈ 35.72, fuel_end ≈ 24.28 > 15 (not enough)
# 250 → daily 500, burn = 47.00, fuel_end = 13.00 ≤ 15 (enough)
# 300 → daily 600, also enough but not minimal.


def _lookahead_dense_routes() -> tuple[LibraryRoute, ...]:
    return (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        _route(2, km=250, frequency=10, typical_station_ids=(10,)),
        _route(3, km=300, frequency=90, typical_station_ids=(10,)),
    )


def test_lookahead_dense_anchors_picks_longer_route() -> None:
    """Fri+Mon fuels with no free weekdays pick the min sufficient daily km."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    result = generate(
        transactions=(
            _tx(friday, qty_liters=40.0, station_id=10),
            _tx(monday, qty_liters=40.0, station_id=10),
        ),
        routes=_lookahead_dense_routes(),
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    friday_wb = _day_by_date(result, friday)
    assert friday_wb.route.km == 250
    assert friday_wb.route.route_id == 2
    assert friday_wb.tank.km == 500
    assert friday_wb.tank.fuel_end <= (TANK - 40.0) + 1e-9
    assert "balance_route" in friday_wb.warnings
    monday_wb = _day_by_date(result, monday)
    assert monday_wb.tank.fuel_start <= (TANK - 40.0) + 1e-9


def test_lookahead_skips_when_burn_needed_non_positive() -> None:
    """burn_needed ≤ 0 → frequency pick (short), do not elongate."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    # Short daily 160 burns ~15 L so Mon stays in corridor; 250 would be waste.
    # fuel_after=40, headroom=45 → burn_needed=-5 ≤ 0
    routes = (
        _route(1, km=80, frequency=100, typical_station_ids=(10,)),
        _route(2, km=250, frequency=10, typical_station_ids=(10,)),
        _route(3, km=190, frequency=50, typical_station_ids=(10,)),
    )
    result = generate(
        transactions=(
            _tx(friday, qty_liters=10.0, station_id=10),
            _tx(monday, qty_liters=10.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=30.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    friday_wb = _day_by_date(result, friday)
    assert friday_wb.route.km == 80
    assert friday_wb.route.route_id == 1
    assert "balance_route" not in friday_wb.warnings


def test_lookahead_rejects_route_over_max_daily_km() -> None:
    """2×km > max_daily_km drops the route even if that km would solve lookahead."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    result = generate(
        transactions=(
            _tx(friday, qty_liters=40.0, station_id=10),
            _tx(monday, qty_liters=40.0, station_id=10),
        ),
        routes=_lookahead_dense_routes(),
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        max_daily_km=400,
    )
    used_km = {d.route.km for d in result.days}
    assert 250 not in used_km
    assert 300 not in used_km
    if result.unsolvable is None:
        friday_wb = _day_by_date(result, friday)
        assert friday_wb.route.km == 190
        assert friday_wb.tank.km <= 400
    else:
        assert result.unsolvable.at_date == monday
        assert result.unsolvable.free_weekdays == 0


def test_lookahead_falls_back_to_full_library_when_group_too_short() -> None:
    """SC-A1: 0 weekdays, typical ≤100 km, library ≥180 → elongate + balance_route."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    routes = (
        _route(1, km=95, frequency=100, typical_station_ids=(10,)),
        _route(2, km=100, frequency=50, typical_station_ids=(10,)),
        _route(3, km=250, frequency=10, typical_station_ids=()),
    )
    result = generate(
        transactions=(
            _tx(friday, qty_liters=40.0, station_id=10),
            _tx(monday, qty_liters=40.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    friday_wb = _day_by_date(result, friday)
    assert friday_wb.route.km == 250
    assert friday_wb.route.route_id == 3
    assert friday_wb.tank.km == 500
    assert "balance_route" in friday_wb.warnings
    monday_wb = _day_by_date(result, monday)
    assert "manual_intervention" not in monday_wb.warnings
    assert result.problematic_days == ()


def test_lookahead_keeps_group_when_typical_already_sufficient() -> None:
    """If the station group already has enough km, do not pick off-group library."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        _route(2, km=250, frequency=10, typical_station_ids=(10,)),
        _route(3, km=300, frequency=90, typical_station_ids=()),
    )
    result = generate(
        transactions=(
            _tx(friday, qty_liters=40.0, station_id=10),
            _tx(monday, qty_liters=40.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    friday_wb = _day_by_date(result, friday)
    assert friday_wb.route.km == 250
    assert friday_wb.route.route_id == 2
    assert friday_wb.route.route_id != 3


def test_lookahead_manual_when_library_cannot_solve() -> None:
    """Neither group nor full library has enough km → v2 manual_intervention."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    routes = (
        _route(1, km=95, frequency=100, typical_station_ids=(10,)),
        _route(2, km=190, frequency=50, typical_station_ids=()),
    )
    result = generate(
        transactions=(
            _tx(friday, qty_liters=40.0, station_id=10),
            _tx(monday, qty_liters=40.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    friday_wb = _day_by_date(result, friday)
    assert friday_wb.route.km == 95
    monday_wb = _day_by_date(result, monday)
    assert "manual_intervention" in monday_wb.warnings
    assert any(p.date == monday and p.reason == "manual_intervention" for p in result.problematic_days)


# ---------------------------------------------------------------------------
# 12. Soft geo sort by direction to the next station (Task 6)
# ---------------------------------------------------------------------------
# Wed 20.05.2026 → Thu 21.05.2026, no free weekdays. Same tank math as §11:
#   km_needed ≈ 478.72; library 250 → daily 500 (sufficient), 190 → 380 (not).
# Points: Yaroslavl 57.63,39.87; Moscow 55.755,37.617; Kostroma 57.766,40.927.

YAROSLAVL = GeoPoint(lat=57.63, lon=39.87)
MOSCOW = GeoPoint(lat=55.755, lon=37.617)
KOSTROMA = GeoPoint(lat=57.766, lon=40.927)
STATION_YAROSLAVL = 31
STATION_MOSCOW = 32
_DIRECTION_DATES = (date(2026, 5, 20), date(2026, 5, 21))


def _direction_station_coords() -> dict[int, GeoPoint]:
    return {
        STATION_YAROSLAVL: YAROSLAVL,
        STATION_MOSCOW: MOSCOW,
    }


def _direction_generate(
    routes: tuple[LibraryRoute, ...],
    *,
    station_coords: dict[int, GeoPoint] | None = None,
    include_coords: bool = True,
) -> GenerateResult:
    day_a, day_b = _DIRECTION_DATES
    coords = None
    if include_coords:
        coords = _direction_station_coords() if station_coords is None else station_coords
    return generate(
        transactions=(
            _tx(day_a, qty_liters=40.0, station_id=STATION_YAROSLAVL),
            _tx(day_b, qty_liters=40.0, station_id=STATION_MOSCOW),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        seed=3,
        station_coords=coords,
    )


def test_direction_yaroslavl_to_moscow_picks_aligned_route() -> None:
    """Same sufficient km: route toward Moscow beats the opposite (Kostroma)."""
    r_moscow = _route(
        12,
        km=250,
        frequency=10,
        typical_station_ids=(STATION_YAROSLAVL,),
        addr_a="Ярославль",
        addr_b="Москва",
        point_a=YAROSLAVL,
        point_b=MOSCOW,
    )
    r_kostroma = _route(
        11,
        km=250,
        frequency=100,
        typical_station_ids=(STATION_YAROSLAVL,),
        addr_a="Ярославль",
        addr_b="Кострома",
        point_a=YAROSLAVL,
        point_b=KOSTROMA,
    )
    result = _direction_generate((r_moscow, r_kostroma))
    assert result.unsolvable is None
    wb = _day_by_date(result, _DIRECTION_DATES[0])
    assert wb.route.route_id == 12
    assert wb.route.addr_b == "Москва"
    assert wb.tank.km == 500


def test_direction_no_station_coords_falls_back_to_km() -> None:
    """No station_coords → min sufficient km / frequency, no crash."""
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(STATION_YAROSLAVL,)),
        _route(2, km=250, frequency=10, typical_station_ids=(STATION_YAROSLAVL,)),
        _route(3, km=300, frequency=90, typical_station_ids=(STATION_YAROSLAVL,)),
    )
    result = _direction_generate(routes, include_coords=False)
    assert result.unsolvable is None
    wb = _day_by_date(result, _DIRECTION_DATES[0])
    assert wb.route.km == 250
    assert wb.route.route_id == 2


def test_direction_no_next_station_coords_falls_back_to_km() -> None:
    """Today coords only / routes without points → fallback to min sufficient km."""
    routes = (
        _route(
            1,
            km=250,
            frequency=10,
            typical_station_ids=(STATION_YAROSLAVL,),
            point_a=YAROSLAVL,
            point_b=MOSCOW,
        ),
        _route(
            2,
            km=250,
            frequency=100,
            typical_station_ids=(STATION_YAROSLAVL,),
            point_a=YAROSLAVL,
            point_b=KOSTROMA,
        ),
        _route(3, km=300, frequency=90, typical_station_ids=(STATION_YAROSLAVL,)),
    )
    result = _direction_generate(
        routes,
        station_coords={STATION_YAROSLAVL: YAROSLAVL},
    )
    assert result.unsolvable is None
    wb = _day_by_date(result, _DIRECTION_DATES[0])
    assert wb.route.km == 250
    assert wb.route.route_id == 2


def test_direction_tank_beats_short_aligned_route() -> None:
    """Aligned short route is dropped when 2×km < km_needed; long hook wins."""
    r_short_moscow = _route(
        21,
        km=190,
        frequency=100,
        typical_station_ids=(STATION_YAROSLAVL,),
        addr_a="Ярославль",
        addr_b="Москва",
        point_a=YAROSLAVL,
        point_b=MOSCOW,
    )
    r_long_hook = _route(
        22,
        km=250,
        frequency=10,
        typical_station_ids=(STATION_YAROSLAVL,),
        addr_a="Ярославль",
        addr_b="Кострома",
        point_a=YAROSLAVL,
        point_b=KOSTROMA,
    )
    result = _direction_generate((r_short_moscow, r_long_hook))
    assert result.unsolvable is None
    wb = _day_by_date(result, _DIRECTION_DATES[0])
    assert wb.route.route_id == 22
    assert wb.route.km == 250
    assert wb.tank.km == 500
    assert wb.tank.fuel_end <= (TANK - 40.0) + 1e-9


# ---------------------------------------------------------------------------
# 13. Partial generation: unsolvable anchor → manual_intervention (Task 7)
# ---------------------------------------------------------------------------


def test_manual_intervention_keeps_later_anchors() -> None:
    """Unsolvable Monday does not abort Friday or Wednesday."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    wednesday = date(2025, 4, 9)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        _route(2, km=150, frequency=40, typical_station_ids=()),
    )
    result = generate(
        transactions=(
            _tx(friday, qty_liters=40.0, station_id=10),
            _tx(monday, qty_liters=40.0, station_id=10),
            _tx(wednesday, qty_liters=40.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert {d.date for d in result.days} >= {friday, monday, wednesday}
    monday_wb = _day_by_date(result, monday)
    assert monday_wb.source == "auto"
    assert "manual_intervention" in monday_wb.warnings
    wednesday_wb = _day_by_date(result, wednesday)
    assert wednesday_wb.date == wednesday
    assert "manual_intervention" not in wednesday_wb.warnings


def test_manual_problematic_days_payload() -> None:
    """problematic_days carries date, reason, detail, fuels, tank_volume."""
    friday = date(2025, 4, 4)
    monday = date(2025, 4, 7)
    q = 40.0
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        _route(2, km=150, frequency=40, typical_station_ids=()),
    )
    result = generate(
        transactions=(
            _tx(friday, qty_liters=40.0, station_id=10),
            _tx(monday, qty_liters=q, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=60_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert len(result.problematic_days) >= 1
    problem = result.problematic_days[0]
    assert isinstance(problem, ProblematicDay)
    assert problem.date == monday
    assert problem.reason == "manual_intervention"
    assert problem.detail
    assert problem.fuel_before > 0
    assert problem.fuel_to_issue == pytest.approx(q)
    assert problem.tank_volume == pytest.approx(TANK)


def test_manual_day_carries_actual_fuel_to_next() -> None:
    """Next day.fuel_start equals the manual day's actual fuel_end (even if > tank)."""
    monday = date(2025, 4, 7)
    wednesday = date(2025, 4, 9)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        *_burn_library(),
    )
    result = generate(
        transactions=(
            _tx(monday, qty_liters=50.0, station_id=10),
            _tx(wednesday, service_type="wash", qty_liters=None, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=50.0,
        odometer_start=10_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    monday_wb = _day_by_date(result, monday)
    wednesday_wb = _day_by_date(result, wednesday)
    assert "manual_intervention" in monday_wb.warnings
    assert monday_wb.tank.fuel_end > TANK
    assert wednesday_wb.tank.fuel_start == pytest.approx(monday_wb.tank.fuel_end, abs=0.01)
    assert wednesday_wb.tank.odometer_start == monday_wb.tank.odometer_end
    assert any(p.date == monday for p in result.problematic_days)


def test_manual_clean_period_has_empty_problematic_days() -> None:
    """Simple solvable period: no problematic_days, unsolvable stays None."""
    day = date(2025, 4, 7)
    routes = (
        _route(1, km=190, frequency=100, typical_station_ids=(10,)),
        *_burn_library(),
    )
    result = generate(
        transactions=(_tx(day, qty_liters=35.0, station_id=10),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2025, 11, 1), "winter"),),
        fuel_start=20.0,
        odometer_start=10_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert result.problematic_days == ()
    assert _day_by_date(result, day).warnings == ()


# ---------------------------------------------------------------------------
# 14. Short burn-in (slice B): pool is max_daily_km, min sufficient km
# ---------------------------------------------------------------------------
# Wed 6 May 2026 → free Thu → Fri fill. TANK=55, NORM=9.4, start=8, Q_Wed=30:
#   typical 95 → daily 190, burn=17.86, fuel_end=20.14
#   Q_Fri=40, headroom=15; need to burn ≥5.14 L on Thursday
#   6 km  → daily 12,  burn=1.13, nxt=19.01 > 15 (not enough)
#   45 km → daily 90,  burn=8.46, nxt=11.68 ≤ 15 and ≥ 0 (min sufficient)
#   80 km → daily 160, burn=15.04, nxt=5.10 ≤ 15 (reaches, but longer)
#   190 km → daily 380, burn=35.72, nxt=-15.58 < 0 (never chosen)


def _short_burn_routes() -> tuple[LibraryRoute, ...]:
    return (
        _route(1, km=95, frequency=100, typical_station_ids=(10,)),
        _route(2, km=135, frequency=50, typical_station_ids=(10,)),
        _route(3, km=45, frequency=8, typical_station_ids=()),
        _route(4, km=6, frequency=90, typical_station_ids=()),
        _route(5, km=190, frequency=80, typical_station_ids=()),
        _route(6, km=80, frequency=70, typical_station_ids=()),
    )


def test_short_burn_weekday_keeps_typical_anchor() -> None:
    """SC-B1: Wed stays typical; Thu short ~45 km; Fri not manual."""
    wednesday = date(2026, 5, 6)
    thursday = date(2026, 5, 7)
    friday = date(2026, 5, 8)
    q_fri = 40.0
    result = generate(
        transactions=(
            _tx(wednesday, qty_liters=30.0, station_id=10),
            _tx(friday, qty_liters=q_fri, station_id=10),
        ),
        routes=_short_burn_routes(),
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=8.0,
        odometer_start=128_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    wed_wb = _day_by_date(result, wednesday)
    assert wed_wb.route.km == 95
    assert wed_wb.route.route_id == 1
    assert "balance_route" not in wed_wb.warnings
    assert "manual_intervention" not in wed_wb.warnings

    thu_wb = _day_by_date(result, thursday)
    assert thu_wb.tank.fuel_issued == pytest.approx(0.0)
    assert thu_wb.route.km == 45
    assert thu_wb.tank.km == 90
    assert thu_wb.tank.fuel_end >= 0.0
    assert thu_wb.tank.fuel_end <= (TANK - q_fri) + 1e-9

    fri_wb = _day_by_date(result, friday)
    assert "manual_intervention" not in fri_wb.warnings
    assert result.problematic_days == ()


def test_burn_that_would_go_negative_is_never_chosen() -> None:
    """A 190 km one-way burn is skipped when remainder is ~20 L."""
    wednesday = date(2026, 5, 6)
    thursday = date(2026, 5, 7)
    friday = date(2026, 5, 8)
    result = generate(
        transactions=(
            _tx(wednesday, qty_liters=30.0, station_id=10),
            _tx(friday, qty_liters=40.0, station_id=10),
        ),
        routes=_short_burn_routes(),
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=8.0,
        odometer_start=128_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    wed_wb = _day_by_date(result, wednesday)
    thu_wb = _day_by_date(result, thursday)
    remaining = wed_wb.tank.fuel_end
    would_burn = burn_for_km(2 * 190, NORM_SUMMER)
    assert would_burn > remaining
    assert thu_wb.route.km != 190
    assert thu_wb.tank.fuel_end >= 0.0


def test_burn_in_headroom_picks_min_sufficient_km_not_frequency() -> None:
    """R4: among burns that already land in headroom, pick smallest daily km."""
    wednesday = date(2026, 5, 6)
    thursday = date(2026, 5, 7)
    friday = date(2026, 5, 8)
    result = generate(
        transactions=(
            _tx(wednesday, qty_liters=30.0, station_id=10),
            _tx(friday, qty_liters=40.0, station_id=10),
        ),
        routes=_short_burn_routes(),
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=8.0,
        odometer_start=128_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    thu_wb = _day_by_date(result, thursday)
    # 45 km (freq 8) and 80 km (freq 70) both reach; min km wins, not frequency.
    assert thu_wb.route.km == 45
    assert thu_wb.route.route_id == 3


def test_burn_max_safe_until_headroom_then_min_sufficient() -> None:
    """Until headroom, max safe burn; the landing day is min sufficient km."""
    monday = date(2026, 4, 6)
    tuesday = date(2026, 4, 7)
    wednesday = date(2026, 4, 8)
    thursday = date(2026, 4, 9)
    routes = (
        _route(1, km=80, frequency=100, typical_station_ids=(10,)),
        _route(2, km=95, frequency=20, typical_station_ids=()),
        _route(3, km=45, frequency=8, typical_station_ids=()),
        _route(4, km=6, frequency=90, typical_station_ids=()),
    )
    result = generate(
        transactions=(
            _tx(monday, qty_liters=40.0, station_id=10),
            _tx(thursday, qty_liters=40.0, station_id=10),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=10.0,
        odometer_start=90_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert result.problematic_days == ()
    tue_wb = _day_by_date(result, tuesday)
    wed_wb = _day_by_date(result, wednesday)
    assert tue_wb.tank.fuel_issued == pytest.approx(0.0)
    assert wed_wb.tank.fuel_issued == pytest.approx(0.0)
    assert tue_wb.route.km == 95
    assert wed_wb.route.km == 45
    assert wed_wb.tank.fuel_end <= (TANK - 40.0) + 1e-9
    assert wed_wb.tank.fuel_end >= 0.0


# ---------------------------------------------------------------------------
# 14. Tank corridor filter + wash prefers short (anchor-corridor spec)
# ---------------------------------------------------------------------------

_PALISADE_TANK = 70.0
_PALISADE_NORM = 14.5
_PALISADE_FUEL_START = 41.13


def test_two_washes_second_day_picks_short_route_in_corridor() -> None:
    """Palisade 03.08+04.08: second wash takes 6 km (12 km daily), tank ~11.84.

    03.08 follows the frequent 95 km typical route (tank → 13.58). On 04.08 the
    95 km round goes negative; corridor keeps 6 km, fuel_end ≈ 11.84, no
    ``manual_intervention``.
    """
    wash_1 = date(2026, 8, 3)  # Monday
    wash_2 = date(2026, 8, 4)  # Tuesday
    station_long = 10
    station_wash = 20
    routes = (
        _route(1, km=6, frequency=5, typical_station_ids=(station_wash,)),
        _route(2, km=95, frequency=39, typical_station_ids=(station_long, station_wash)),
    )
    result = generate(
        transactions=(
            _tx(wash_1, service_type="wash", qty_liters=None, station_id=station_long),
            _tx(wash_2, service_type="wash", qty_liters=None, station_id=station_wash),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=_PALISADE_TANK,
        norm_summer=_PALISADE_NORM,
        norm_winter=_PALISADE_NORM,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=_PALISADE_FUEL_START,
        odometer_start=80_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    day2 = _day_by_date(result, wash_2)
    assert day2.route.km == 6
    assert day2.tank.km == 12
    assert day2.tank.fuel_end == pytest.approx(11.84, abs=0.01)
    assert day2.tank.fuel_end >= 0.0
    assert "manual_intervention" not in day2.warnings
    assert all(p.date != wash_2 for p in result.problematic_days)


def test_wash_lookahead_replaces_short_route_when_next_refill_overflows() -> None:
    """Wash prefers min-km, but lookahead lengthens when Q_next would overflow.

    Wash 04.08 at 60 L, next fill 05.08 +50 L, tank 70 L, norm 14.5:
    short 6 km would leave ~58.26; 58.26+50 exceeds 70, so the day is replaced
    by a route long enough for the existing lookahead burn-off.
    """
    wash_day = date(2026, 8, 4)  # Tuesday
    fuel_day = date(2026, 8, 5)  # Wednesday — no free weekday between
    station = 10
    # Lookahead km_needed uses fuel_before+Q (40 L / 14.5 × 100 ≈ 275.86 daily),
    # so the long typical must be ≥ 138 km shoulder. 140 km daily 280.
    routes = (
        _route(1, km=6, frequency=5, typical_station_ids=(station,)),
        _route(2, km=140, frequency=39, typical_station_ids=(station,)),
    )
    result = generate(
        transactions=(
            _tx(wash_day, service_type="wash", qty_liters=None, station_id=station),
            _tx(fuel_day, qty_liters=50.0, station_id=station),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=_PALISADE_TANK,
        norm_summer=_PALISADE_NORM,
        norm_winter=_PALISADE_NORM,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=60.0,
        odometer_start=80_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    wash_wb = _day_by_date(result, wash_day)
    assert wash_wb.route.km == 140
    assert wash_wb.route.route_id == 2
    assert wash_wb.tank.km == 280
    assert "balance_route" in wash_wb.warnings
    assert "manual_intervention" not in wash_wb.warnings
    fuel_wb = _day_by_date(result, fuel_day)
    assert fuel_wb.tank.fuel_start + 50.0 <= _PALISADE_TANK + 1e-9
    assert "manual_intervention" not in fuel_wb.warnings


# ---------------------------------------------------------------------------
# T1: fleet pool helpers — city key, home base, vehicle_id / own_vehicle_id
# ---------------------------------------------------------------------------


def test_city_key_sergiev_posad_matches_compound_name() -> None:
    from_addr = _city_key("г.Сергиев Посад, ул.Маслиева, д.1")
    from_city = _city_key("Сергиев Посад")
    assert from_addr == from_city
    assert from_addr == "сергиев посад"


def test_city_key_compound_and_g_prefix_cities() -> None:
    assert _city_key("г.Переславль-Залесский, ул.Ленина") == "переславль залесский"
    assert _city_key("Переславль Залесский") == "переславль залесский"
    assert _city_key("г.Вологда, ул.Мира") == "вологда"
    assert _city_key("г. Мантурово, ул.Красная") == "мантурово"
    assert _city_key("г.Нижний Новгород, ул.Горького") == "нижний новгород"


def test_is_home_base_kuznetskaya() -> None:
    assert _is_home_base("ул. Кузнецкая, д.18Б") is True
    assert _is_home_base("ул.Кузнецкая, д.18Б") is True
    assert _is_home_base("улица Кузнецкая, дом 18Б") is True
    assert _is_home_base("ул.Кузнецкая") is True
    assert _is_home_base("г.Вологда, ул.Мира, д.10") is False


def test_library_route_vehicle_id_defaults_to_zero() -> None:
    route = LibraryRoute(
        route_id=1,
        addr_a="A",
        addr_b="B",
        km=10,
        frequency=1,
        typical_station_ids=(),
    )
    assert route.vehicle_id == 0


def test_generate_own_vehicle_id_defaults_to_zero() -> None:
    param = inspect.signature(generate).parameters["own_vehicle_id"]
    assert param.default == 0


def test_generate_without_own_vehicle_id_treats_routes_as_own() -> None:
    """Legacy fixtures omit vehicle_id; default 0 keeps every route 'own'."""
    fuel_day = date(2025, 4, 7)
    result = generate(
        transactions=(_tx(fuel_day),),
        routes=(_route(1, km=100, typical_station_ids=(10,)),),
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=(),
        fuel_start=40.0,
        odometer_start=10_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
    )
    assert result.unsolvable is None
    assert len(result.days) == 1
    assert result.days[0].route.route_id == 1


# ---------------------------------------------------------------------------
# T2: fleet-pool corridor cascade, own-vs-foreign rank, persist without donor id
# ---------------------------------------------------------------------------

_HOME = "Кострома, ул. Кузнецкая, д.18Б"
_SERGIEV = "г.Сергиев Посад, ул.Маслиева, д.1"
_VOLOGDA = "г.Вологда, ул.Мира"
_MANTUROVO = "г. Мантурово, ул.Красная"


def test_wash_typical_too_long_picks_own_short_from_fleet() -> None:
    """Wash: typical group only 95 km (does not fit), own 6 km in fleet → 12 km.

    Monjaro-shaped tank: 95 km round-trip would go negative; fleet short route
    keeps the day in corridor without ``manual_intervention``.
    """
    wash_day = date(2026, 7, 1)  # Wednesday
    station = 10
    own_id = 2
    routes = (
        _route(
            1,
            km=95,
            frequency=40,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_MANTUROVO,
            vehicle_id=own_id,
        ),
        _route(
            2,
            km=6,
            frequency=5,
            typical_station_ids=(),
            addr_a=_HOME,
            addr_b=_HOME,
            vehicle_id=own_id,
        ),
    )
    result = generate(
        transactions=(_tx(wash_day, service_type="wash", qty_liters=None, station_id=station),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=60.0,
        norm_summer=9.5,
        norm_winter=9.5,
        season_switches=(),
        fuel_start=7.21,
        odometer_start=50_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    day = _day_by_date(result, wash_day)
    assert day.route.km == 6
    assert day.tank.km == 12
    assert day.tank.fuel_end >= 0.0
    assert "manual_intervention" not in day.warnings
    assert all(p.date != wash_day for p in result.problematic_days)


def test_own_long_outside_corridor_borrows_foreign_kuznetskaya() -> None:
    """Own 280 is outside corridor; foreign 265 Kuznetskaya fits → 530 km, borrowed."""
    fuel_day = date(2026, 7, 27)  # Monday
    station = 10
    own_id = 4
    donor_id = 3
    routes = (
        _route(
            40,
            km=280,
            frequency=80,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_VOLOGDA,
            vehicle_id=own_id,
        ),
        _route(
            30,
            km=265,
            frequency=20,
            typical_station_ids=(),
            addr_a=_HOME,
            addr_b=_SERGIEV,
            vehicle_id=donor_id,
        ),
    )
    result = generate(
        transactions=(_tx(fuel_day, qty_liters=50.0, station_id=station),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=9.4,
        season_switches=(),
        fuel_start=1.43,
        odometer_start=90_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    day = _day_by_date(result, fuel_day)
    assert day.route.km == 265
    assert day.tank.km == 530
    assert day.tank.fuel_end >= 0.0
    assert "borrowed_route" in day.warnings
    assert "manual_intervention" not in day.warnings
    assert day.route.route_id is None
    assert all(leg.route_id is None for leg in day.legs)
    assert day.legs[0].km == 265
    assert day.legs[0].addr_a == _HOME
    assert day.legs[0].addr_b == _SERGIEV


def test_own_and_foreign_same_km_picks_own_route_id() -> None:
    """Own 6 km and foreign 6 km both in corridor → own wins, route_id persisted."""
    wash_day = date(2026, 7, 2)  # Thursday
    station = 10
    own_id = 2
    own_route_id = 20
    foreign_route_id = 1  # lower id so min(route_id) would wrongly pick donor
    routes = (
        _route(
            foreign_route_id,
            km=6,
            frequency=90,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_HOME,
            vehicle_id=1,
        ),
        _route(
            own_route_id,
            km=6,
            frequency=5,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_HOME,
            vehicle_id=own_id,
        ),
    )
    result = generate(
        transactions=(_tx(wash_day, service_type="wash", qty_liters=None, station_id=station),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=60.0,
        norm_summer=9.5,
        norm_winter=9.5,
        season_switches=(),
        fuel_start=20.0,
        odometer_start=50_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    day = _day_by_date(result, wash_day)
    assert day.route.km == 6
    assert day.route.route_id == own_route_id
    assert day.route.route_id is not None
    assert all(leg.route_id == own_route_id for leg in day.legs)
    assert "borrowed_route" not in day.warnings


def test_lookahead_does_not_pick_280_when_265_in_corridor() -> None:
    """Lookahead window ~494–547: 265 (530) is in S; must not pick 280 (560)."""
    friday = date(2026, 7, 24)
    monday = date(2026, 7, 27)
    station = 10
    own_id = 4
    routes = (
        _route(
            40,
            km=280,
            frequency=80,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_VOLOGDA,
            vehicle_id=own_id,
        ),
        _route(
            30,
            km=265,
            frequency=20,
            typical_station_ids=(),
            addr_a=_HOME,
            addr_b=_SERGIEV,
            vehicle_id=3,
        ),
    )
    result = generate(
        transactions=(
            _tx(friday, qty_liters=50.0, station_id=station),
            _tx(monday, qty_liters=50.0, station_id=station),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=9.4,
        season_switches=(),
        fuel_start=1.43,
        odometer_start=90_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    friday_wb = _day_by_date(result, friday)
    assert friday_wb.route.km == 265
    assert friday_wb.tank.km == 530
    assert friday_wb.tank.km != 560
    assert friday_wb.tank.fuel_end >= 0.0
    assert "borrowed_route" in friday_wb.warnings
    assert "manual_intervention" not in friday_wb.warnings


def test_burn_in_borrows_foreign_short_when_own_outside_corridor() -> None:
    """Burn-in: own 95 km goes negative; foreign 6 km in corridor → 12 km, borrowed."""
    wednesday = date(2026, 5, 6)
    thursday = date(2026, 5, 7)
    friday = date(2026, 5, 8)
    station = 10
    own_id = 2
    donor_route_id = 6
    routes = (
        _route(
            1,
            km=95,
            frequency=100,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_MANTUROVO,
            vehicle_id=own_id,
        ),
        _route(
            donor_route_id,
            km=6,
            frequency=90,
            typical_station_ids=(),
            addr_a=_HOME,
            addr_b=_HOME,
            vehicle_id=3,
        ),
    )
    result = generate(
        transactions=(
            _tx(wednesday, qty_liters=30.0, station_id=station),
            _tx(friday, qty_liters=40.0, station_id=station),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=3.0,
        odometer_start=128_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    thu_wb = _day_by_date(result, thursday)
    assert thu_wb.tank.fuel_issued == pytest.approx(0.0)
    assert thu_wb.route.km == 6
    assert thu_wb.tank.km == 12
    assert thu_wb.tank.fuel_end >= 0.0
    assert "borrowed_route" in thu_wb.warnings
    assert thu_wb.route.route_id is None
    assert all(leg.route_id is None for leg in thu_wb.legs)


# ---------------------------------------------------------------------------
# Home-oriented round-trip: twin lookup + emit orientation
# ---------------------------------------------------------------------------

_VLADIMIR = "г.Владимир, ул.Добросельская"


def test_find_home_twin_swapped_endpoints_same_vehicle_km() -> None:
    chosen = _route(59, km=225, addr_a=_VLADIMIR, addr_b=_HOME, vehicle_id=4)
    twin = _route(64, km=225, addr_a=_HOME, addr_b=_VLADIMIR, vehicle_id=4)
    catalog = (
        chosen,
        twin,
        _route(70, km=225, addr_a=_HOME, addr_b=_VLADIMIR, vehicle_id=3),
        _route(80, km=280, addr_a=_HOME, addr_b=_VLADIMIR, vehicle_id=4),
    )
    found = _find_home_twin(chosen, catalog)
    assert found is twin


def test_find_home_twin_picks_min_route_id() -> None:
    chosen = _route(59, km=225, addr_a=_VLADIMIR, addr_b=_HOME, vehicle_id=4)
    catalog = (
        chosen,
        _route(80, km=225, addr_a=_HOME, addr_b=_VLADIMIR, vehicle_id=4),
        _route(64, km=225, addr_a=_HOME, addr_b=_VLADIMIR, vehicle_id=4),
    )
    found = _find_home_twin(chosen, catalog)
    assert found is not None
    assert found.route_id == 64


def test_find_home_twin_does_not_return_chosen() -> None:
    chosen = _route(59, km=225, addr_a=_VLADIMIR, addr_b=_HOME, vehicle_id=4)
    assert _find_home_twin(chosen, (chosen,)) is None


def test_orient_home_round_trip_object_first_with_twin() -> None:
    chosen = _route(59, km=225, addr_a=_VLADIMIR, addr_b=_HOME, vehicle_id=4)
    twin = _route(
        64,
        km=225,
        addr_a=_HOME,
        addr_b="г.владимир, ул.Добросельская",
        vehicle_id=4,
    )
    oriented = _orient_home_round_trip(chosen, catalog=(chosen, twin), own_vehicle_id=4)
    assert _is_home_base(oriented.addr_a)
    assert oriented.addr_b == _VLADIMIR
    assert oriented.route_id == 64
    assert oriented.km == 225
    assert oriented.vehicle_id == 4


def test_orient_home_round_trip_object_first_without_twin() -> None:
    chosen = _route(59, km=225, addr_a=_VLADIMIR, addr_b=_HOME, vehicle_id=4)
    oriented = _orient_home_round_trip(chosen, catalog=(chosen,), own_vehicle_id=4)
    assert _is_home_base(oriented.addr_a)
    assert oriented.addr_b == _VLADIMIR
    assert oriented.route_id == 59


def test_orient_home_round_trip_borrowed_keeps_library_route_id() -> None:
    chosen = _route(30, km=265, addr_a=_SERGIEV, addr_b=_HOME, vehicle_id=3)
    twin = _route(31, km=265, addr_a=_HOME, addr_b=_SERGIEV, vehicle_id=3)
    oriented = _orient_home_round_trip(
        chosen, catalog=(chosen, twin), own_vehicle_id=4
    )
    assert _is_home_base(oriented.addr_a)
    assert oriented.addr_b == _SERGIEV
    assert oriented.route_id == 30
    assert oriented.vehicle_id == 3


def test_orient_home_round_trip_home_first_is_identity() -> None:
    chosen = _route(64, km=225, addr_a=_HOME, addr_b=_VLADIMIR, vehicle_id=4)
    twin = _route(59, km=225, addr_a=_VLADIMIR, addr_b=_HOME, vehicle_id=4)
    oriented = _orient_home_round_trip(chosen, catalog=(chosen, twin), own_vehicle_id=4)
    assert oriented is chosen


def test_orient_home_round_trip_both_home_is_identity() -> None:
    chosen = _route(2, km=6, addr_a=_HOME, addr_b=_HOME, vehicle_id=2)
    assert _orient_home_round_trip(chosen, catalog=(chosen,), own_vehicle_id=2) is chosen


def test_orient_home_round_trip_neither_home_is_identity() -> None:
    chosen = _route(1, km=190, addr_a="A", addr_b="B")
    assert _orient_home_round_trip(chosen, catalog=(chosen,), own_vehicle_id=0) is chosen


def test_generate_orients_own_object_first_to_twin_route_id() -> None:
    fuel_day = date(2026, 7, 13)  # Monday
    station = 10
    own_id = 4
    routes = (
        _route(
            59,
            km=225,
            frequency=80,
            typical_station_ids=(station,),
            addr_a=_VLADIMIR,
            addr_b=_HOME,
            vehicle_id=own_id,
        ),
        _route(
            64,
            km=225,
            frequency=20,
            typical_station_ids=(),
            addr_a=_HOME,
            addr_b=_VLADIMIR,
            vehicle_id=own_id,
        ),
    )
    result = generate(
        transactions=(_tx(fuel_day, qty_liters=40.0, station_id=station),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=(),
        fuel_start=20.0,
        odometer_start=10_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    day = _day_by_date(result, fuel_day)
    assert day.route.km == 225
    assert day.tank.km == 450
    assert _is_home_base(day.legs[0].addr_a)
    assert day.legs[0].addr_b == _VLADIMIR
    assert day.legs[0].addr_a == _HOME
    assert day.route.route_id == 64
    assert all(leg.route_id == 64 for leg in day.legs)
    assert "borrowed_route" not in day.warnings
    assert day.tank.fuel_end == pytest.approx(
        20.0 + 40.0 - burn_for_km(450, NORM_SUMMER), abs=0.01
    )


def test_generate_orients_object_first_without_twin_keeps_route_id() -> None:
    fuel_day = date(2026, 7, 13)
    own_id = 4
    routes = (
        _route(
            59,
            km=225,
            frequency=80,
            typical_station_ids=(10,),
            addr_a=_VLADIMIR,
            addr_b=_HOME,
            vehicle_id=own_id,
        ),
    )
    result = generate(
        transactions=(_tx(fuel_day, qty_liters=40.0, station_id=10),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=(),
        fuel_start=20.0,
        odometer_start=10_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    day = _day_by_date(result, fuel_day)
    assert _is_home_base(day.legs[0].addr_a)
    assert day.legs[0].addr_b == _VLADIMIR
    assert day.route.route_id == 59
    assert all(leg.route_id == 59 for leg in day.legs)


def test_generate_orients_borrowed_object_first_route_id_none() -> None:
    fuel_day = date(2026, 7, 27)  # Monday
    station = 10
    own_id = 4
    donor_id = 3
    routes = (
        _route(
            40,
            km=280,
            frequency=80,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_VOLOGDA,
            vehicle_id=own_id,
        ),
        _route(
            30,
            km=265,
            frequency=20,
            typical_station_ids=(),
            addr_a=_SERGIEV,
            addr_b=_HOME,
            vehicle_id=donor_id,
        ),
        _route(
            31,
            km=265,
            frequency=10,
            typical_station_ids=(),
            addr_a=_HOME,
            addr_b=_SERGIEV,
            vehicle_id=donor_id,
        ),
    )
    result = generate(
        transactions=(_tx(fuel_day, qty_liters=50.0, station_id=station),),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=55.0,
        norm_summer=9.4,
        norm_winter=9.4,
        season_switches=(),
        fuel_start=1.43,
        odometer_start=90_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    day = _day_by_date(result, fuel_day)
    assert day.route.km == 265
    assert day.tank.km == 530
    assert "borrowed_route" in day.warnings
    assert day.route.route_id is None
    assert all(leg.route_id is None for leg in day.legs)
    assert _is_home_base(day.legs[0].addr_a)
    assert day.legs[0].addr_a == _HOME
    assert day.legs[0].addr_b == _SERGIEV


def test_generate_burn_in_orients_object_first_from_home() -> None:
    wednesday = date(2026, 5, 6)
    thursday = date(2026, 5, 7)
    friday = date(2026, 5, 8)
    station = 10
    own_id = 2
    routes = (
        _route(
            1,
            km=95,
            frequency=100,
            typical_station_ids=(station,),
            addr_a=_HOME,
            addr_b=_MANTUROVO,
            vehicle_id=own_id,
        ),
        _route(
            7,
            km=6,
            frequency=90,
            typical_station_ids=(),
            addr_a=_VOLOGDA,
            addr_b=_HOME,
            vehicle_id=own_id,
        ),
    )
    result = generate(
        transactions=(
            _tx(wednesday, qty_liters=30.0, station_id=station),
            _tx(friday, qty_liters=40.0, station_id=station),
        ),
        routes=routes,
        hooks={},
        driver_id=DRIVER_ID,
        tank_volume_liters=TANK,
        norm_summer=NORM_SUMMER,
        norm_winter=NORM_WINTER,
        season_switches=((date(2026, 11, 1), "winter"),),
        fuel_start=3.0,
        odometer_start=128_000,
        holidays=frozenset(),
        extra_workdays=frozenset(),
        own_vehicle_id=own_id,
    )
    assert result.unsolvable is None
    thu_wb = _day_by_date(result, thursday)
    assert thu_wb.tank.fuel_issued == pytest.approx(0.0)
    assert thu_wb.route.km == 6
    assert thu_wb.tank.km == 12
    assert _is_home_base(thu_wb.legs[0].addr_a)
    assert thu_wb.legs[0].addr_a == _HOME
    assert thu_wb.legs[0].addr_b == _VOLOGDA
    assert thu_wb.route.route_id == 7

