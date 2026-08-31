"""Pure waybill-period generator: anchors, burn-in, route pick (no I/O)."""

from __future__ import annotations

import re
from collections import defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from datetime import date, timedelta
from random import Random

from core.gsm.balance import BalanceViolation, apply_day, burn_for_km
from core.gsm.geo import GeoPoint, angle_diff_deg, bearing_deg
from core.gsm.models import LegPlan, RouteRef, TankState, Transaction, WaybillDay
from core.gsm.season import SeasonSwitch, norm_for
from core.work_calendar import is_working_day

_ROUND_TRIP_LEGS = 2
_ALIGNED_MAX_ANGLE_DEG = 90.0
_ANCHOR_SERVICES = frozenset({"fuel", "wash"})
_MANUAL_REASON = "manual_intervention"

_NextFuel = tuple[date, float, tuple[int, ...]]


@dataclass(frozen=True, slots=True)
class LibraryRoute:
    """Frozen library entry used for anchor and burn-in route selection."""

    route_id: int
    addr_a: str
    addr_b: str
    km: int
    frequency: int
    typical_station_ids: tuple[int, ...]
    vehicle_id: int = 0
    point_a: GeoPoint | None = None
    point_b: GeoPoint | None = None


@dataclass(frozen=True, slots=True)
class UnsolvableInfo:
    """Typed failure when headroom / corridor cannot be satisfied."""

    reason: str
    at_date: date
    fuel_before: float
    fuel_to_issue: float
    tank_volume: float
    free_weekdays: int
    detail: str


@dataclass(frozen=True, slots=True)
class ProblematicDay:
    """Anchor that could not stay in the tank corridor / headroom."""

    date: date
    reason: str
    detail: str
    fuel_before: float
    fuel_to_issue: float
    tank_volume: float


@dataclass(frozen=True, slots=True)
class GenerateResult:
    """Frozen period generation result."""

    days: tuple[WaybillDay, ...]
    unsolvable: UnsolvableInfo | None
    warnings: tuple[str, ...]
    problematic_days: tuple[ProblematicDay, ...] = ()


def _norm_addr(s: str) -> str:
    t = s.casefold().replace("ё", "е")
    t = t.replace("улица", "ул").replace("ул.", "ул")
    return " ".join(t.split())


def _is_home_base(addr: str) -> bool:
    n = _norm_addr(addr)
    if "кузнецкая" not in n:
        return False
    if "18" in n:
        return True
    # T1: «ул.Кузнецкая» without a house number is the same home base.
    return re.search(r"\d", n) is None


def _city_key(addr: str) -> str:
    n = _norm_addr(addr)
    for name in (
        "сергиев посад",
        "переславль-залесский",
        "переславль залесский",
        "нижний новгород",
        "н.новгород",
    ):
        if name in n:
            return name.replace("-", " ")
    match = re.search(r"г\.\s*([^,]+)", n)
    if match:
        return match.group(1).strip()
    return n[:40]


def _home_base_routes(routes: Sequence[LibraryRoute]) -> tuple[LibraryRoute, ...]:
    return tuple(
        route
        for route in routes
        if _is_home_base(route.addr_a) or _is_home_base(route.addr_b)
    )


def _fleet_pool(routes: Sequence[LibraryRoute]) -> tuple[LibraryRoute, ...]:
    """Home-base routes when present; else the full capped list (legacy A/B)."""
    home = _home_base_routes(routes)
    return home if home else tuple(routes)


def generate(
    *,
    transactions: Sequence[Transaction],
    routes: Sequence[LibraryRoute],
    hooks: Mapping[tuple[int, int], float],
    driver_id: int,
    tank_volume_liters: float,
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    fuel_start: float,
    odometer_start: int,
    hook_threshold_km: float = 13.0,
    holidays: frozenset[date] | set[date] = frozenset(),
    extra_workdays: frozenset[date] | set[date] = frozenset(),
    seed: int = 0,
    max_daily_km: int = 700,
    station_coords: Mapping[int, GeoPoint] | None = None,
    own_vehicle_id: int = 0,
) -> GenerateResult:
    """Build draft waybill days for a transaction period (pure, deterministic).

    ``season_switches`` selects summer/winter norm per day; empty → summer.
    ``own_vehicle_id`` defaults to 0 so legacy fixtures without ``vehicle_id``
    treat every library route as own. The burn-in pool is the loaded fleet
    (home-base routes when present), not sliced to this vehicle.
    """
    holiday_set = frozenset(holidays)
    extra_set = frozenset(extra_workdays)
    switches = tuple(sorted(season_switches, key=lambda s: s[0]))
    anchors = _build_anchors(transactions)
    if not anchors:
        return GenerateResult(days=(), unsolvable=None, warnings=())

    route_list = _routes_within_daily_cap(tuple(routes), max_daily_km)
    home_pool = _home_base_routes(route_list)
    # Fleet∩home for burn-in; do not slice to own_vehicle_id. Empty home
    # (legacy A/B fixtures, all vehicle_id=0) keeps the full capped list.
    burn_source = home_pool if home_pool else route_list
    burn_routes = _ordered_burn_routes(burn_source, seed=seed)
    occupied_dates = {day for day, _txs in anchors}
    period_warnings: list[str] = []
    waybills: list[WaybillDay] = []
    problems: list[ProblematicDay] = []
    fuel = float(fuel_start)
    odometer = int(odometer_start)
    prev_date: date | None = None

    for idx, (anchor_date, txs) in enumerate(anchors):
        station_ids = _station_ids(txs)
        fuel_issued = _fuel_issued(txs)
        free_days = (
            _free_weekdays_between(prev_date, anchor_date, holiday_set, extra_set)
            if prev_date is not None
            else ()
        )

        headroom_fail: UnsolvableInfo | None = None
        if prev_date is not None and fuel_issued > 0:
            target_max = tank_volume_liters - fuel_issued
            burn_plan, _fuel_after, unsolvable = _plan_burn_in(
                fuel=fuel,
                free_days=free_days,
                burn_routes=burn_routes,
                target_fuel_max=target_max,
                tank_volume_liters=tank_volume_liters,
                norm_summer=norm_summer,
                norm_winter=norm_winter,
                season_switches=switches,
                at_date=anchor_date,
                fuel_to_issue=fuel_issued,
                own_vehicle_id=own_vehicle_id,
            )
            if unsolvable is not None:
                headroom_fail = unsolvable
            else:
                for burn_day, burn_route in burn_plan:
                    burn_warnings: tuple[str, ...] = (
                        ("borrowed_route",)
                        if burn_route.vehicle_id != own_vehicle_id
                        else ()
                    )
                    wb, fuel, odometer, problem = _emit_maybe_manual(
                        day=burn_day,
                        route=burn_route,
                        fuel_start=fuel,
                        fuel_issued=0.0,
                        odometer_start=odometer,
                        driver_id=driver_id,
                        tank_volume_liters=tank_volume_liters,
                        norm_summer=norm_summer,
                        norm_winter=norm_winter,
                        season_switches=switches,
                        day_warnings=burn_warnings,
                        detail=f"burn-in day {burn_day.isoformat()} left corridor",
                        own_vehicle_id=own_vehicle_id,
                        catalog=route_list,
                    )
                    waybills.append(wb)
                    if problem is not None:
                        problems.append(problem)
                        period_warnings.append(_MANUAL_REASON)

        day_warnings: list[str] = []
        if not is_working_day(anchor_date, holidays=holiday_set, extra_workdays=extra_set):
            day_warnings.append("weekend_anchor")
            period_warnings.append("weekend_anchor")

        next_fuel = _next_fuel_anchor(anchors, idx)
        free_until_next = _free_weekdays_until(
            start_exclusive=anchor_date,
            end_exclusive=next_fuel[0] if next_fuel is not None else None,
            occupied=occupied_dates,
            holidays=holiday_set,
            extra_workdays=extra_set,
        )
        route, hook_warnings = _select_anchor_route_lookahead(
            station_ids=station_ids,
            routes=route_list,
            hooks=hooks,
            hook_threshold_km=hook_threshold_km,
            fuel_before=fuel,
            q_today=fuel_issued,
            next_fuel=next_fuel,
            tank_volume_liters=tank_volume_liters,
            norm=norm_for(
                anchor_date,
                norm_summer=norm_summer,
                norm_winter=norm_winter,
                switches=switches,
            ),
            free_until_next=free_until_next,
            burn_routes=burn_routes,
            norm_summer=norm_summer,
            norm_winter=norm_winter,
            season_switches=switches,
            seed=seed,
            station_coords=station_coords,
            own_vehicle_id=own_vehicle_id,
        )
        for code in hook_warnings:
            day_warnings.append(code)
            period_warnings.append(code)

        if headroom_fail is not None:
            day_warnings.append(_MANUAL_REASON)
            period_warnings.append(_MANUAL_REASON)
        detail = (
            headroom_fail.detail
            if headroom_fail is not None
            else f"anchor day {anchor_date.isoformat()} left corridor"
        )
        wb, fuel, odometer, problem = _emit_maybe_manual(
            day=anchor_date,
            route=route,
            fuel_start=fuel,
            fuel_issued=fuel_issued,
            odometer_start=odometer,
            driver_id=driver_id,
            tank_volume_liters=tank_volume_liters,
            norm_summer=norm_summer,
            norm_winter=norm_winter,
            season_switches=switches,
            day_warnings=tuple(day_warnings),
            detail=detail,
            force_manual=headroom_fail is not None,
            own_vehicle_id=own_vehicle_id,
            catalog=route_list,
        )
        waybills.append(wb)
        if problem is not None:
            problems.append(problem)
            period_warnings.append(_MANUAL_REASON)
        prev_date = anchor_date

    return GenerateResult(
        days=tuple(waybills),
        unsolvable=None,
        warnings=_dedupe_warnings(period_warnings),
        problematic_days=tuple(problems),
    )


def _build_anchors(
    transactions: Sequence[Transaction],
) -> list[tuple[date, tuple[Transaction, ...]]]:
    by_date: dict[date, list[Transaction]] = defaultdict(list)
    for tx in sorted(transactions, key=lambda t: (t.ts, t.station_id or 0)):
        if tx.service_type not in _ANCHOR_SERVICES:
            continue
        by_date[tx.ts.date()].append(tx)
    return [(d, tuple(by_date[d])) for d in sorted(by_date)]


def _station_ids(txs: Sequence[Transaction]) -> tuple[int, ...]:
    seen: list[int] = []
    for tx in txs:
        if tx.station_id is None:
            continue
        if tx.station_id not in seen:
            seen.append(tx.station_id)
    return tuple(seen)


def _fuel_issued(txs: Sequence[Transaction]) -> float:
    total = 0.0
    for tx in txs:
        if tx.service_type == "fuel" and tx.qty_liters is not None:
            total += float(tx.qty_liters)
    return total


def _find_home_twin(
    chosen: LibraryRoute, catalog: Sequence[LibraryRoute]
) -> LibraryRoute | None:
    twins = [
        candidate
        for candidate in catalog
        if candidate.route_id != chosen.route_id
        and candidate.vehicle_id == chosen.vehicle_id
        and candidate.km == chosen.km
        and _norm_addr(candidate.addr_a) == _norm_addr(chosen.addr_b)
        and _norm_addr(candidate.addr_b) == _norm_addr(chosen.addr_a)
    ]
    if not twins:
        return None
    return min(twins, key=lambda route: route.route_id)


def _orient_home_round_trip(
    chosen: LibraryRoute,
    *,
    catalog: Sequence[LibraryRoute],
    own_vehicle_id: int,
) -> LibraryRoute:
    a_home = _is_home_base(chosen.addr_a)
    b_home = _is_home_base(chosen.addr_b)
    if a_home == b_home or a_home:
        return chosen
    twin = _find_home_twin(chosen, catalog)
    persist_id = (
        twin.route_id
        if twin is not None and chosen.vehicle_id == own_vehicle_id
        else chosen.route_id
    )
    return replace(
        chosen,
        addr_a=chosen.addr_b,
        addr_b=chosen.addr_a,
        route_id=persist_id,
    )


def _persisted_route_id(route: LibraryRoute, own_vehicle_id: int) -> int | None:
    if route.vehicle_id != own_vehicle_id:
        return None
    return route.route_id


def _to_route_ref(route: LibraryRoute, *, own_vehicle_id: int = 0) -> RouteRef:
    return RouteRef(
        route_id=_persisted_route_id(route, own_vehicle_id),
        addr_a=route.addr_a,
        addr_b=route.addr_b,
        km=route.km,
    )


def _daily_km(route: LibraryRoute) -> int:
    """Round-trip distance: two library legs."""
    return _ROUND_TRIP_LEGS * route.km


def _routes_within_daily_cap(
    routes: Sequence[LibraryRoute], max_daily_km: int
) -> tuple[LibraryRoute, ...]:
    return tuple(r for r in routes if _daily_km(r) <= max_daily_km)


def _round_trip_legs(
    route: LibraryRoute, *, own_vehicle_id: int = 0
) -> tuple[LegPlan, LegPlan]:
    route_id = _persisted_route_id(route, own_vehicle_id)
    return (
        LegPlan(
            route_id=route_id,
            addr_a=route.addr_a,
            addr_b=route.addr_b,
            km=route.km,
        ),
        LegPlan(
            route_id=route_id,
            addr_a=route.addr_b,
            addr_b=route.addr_a,
            km=route.km,
        ),
    )


def _ordered_burn_routes(routes: Sequence[LibraryRoute], *, seed: int) -> tuple[LibraryRoute, ...]:
    """Order burn candidates already within max_daily_km (frequency, then id)."""
    rng = Random(seed)
    # frequency desc; ties → lower route_id; seed as final stable key
    tie = {r.route_id: rng.random() for r in routes}
    ordered = sorted(routes, key=lambda r: (-r.frequency, r.route_id, tie[r.route_id]))
    return tuple(ordered)


def _free_weekdays_between(
    start_exclusive: date,
    end_exclusive: date,
    holidays: frozenset[date],
    extra_workdays: frozenset[date],
) -> tuple[date, ...]:
    days: list[date] = []
    current = start_exclusive + timedelta(days=1)
    while current < end_exclusive:
        if is_working_day(current, holidays=holidays, extra_workdays=extra_workdays):
            days.append(current)
        current += timedelta(days=1)
    return tuple(days)


def _next_fuel_anchor(
    anchors: Sequence[tuple[date, tuple[Transaction, ...]]],
    current_index: int,
) -> _NextFuel | None:
    """Next anchor with fuel_issued > 0 (wash without liters is not Q_next)."""
    for later_date, later_txs in anchors[current_index + 1 :]:
        issued = _fuel_issued(later_txs)
        if issued > 0:
            return later_date, issued, _station_ids(later_txs)
    return None


def _free_weekdays_until(
    *,
    start_exclusive: date,
    end_exclusive: date | None,
    occupied: set[date],
    holidays: frozenset[date],
    extra_workdays: frozenset[date],
) -> tuple[date, ...]:
    if end_exclusive is None:
        return ()
    raw = _free_weekdays_between(start_exclusive, end_exclusive, holidays, extra_workdays)
    return tuple(day for day in raw if day not in occupied)


def _anchor_route_group(
    *,
    station_ids: Sequence[int],
    routes: Sequence[LibraryRoute],
    hooks: Mapping[tuple[int, int], float],
) -> tuple[tuple[LibraryRoute, ...], str, dict[int, float]]:
    """Candidate group as today: typical_station / hook / all."""
    if not routes:
        raise ValueError("routes must not be empty")

    if station_ids:
        matching = [
            r
            for r in routes
            if any(sid in r.typical_station_ids for sid in station_ids)
        ]
        covering_all = [
            r
            for r in matching
            if all(sid in r.typical_station_ids for sid in station_ids)
        ]
        typical = covering_all or matching
        if typical:
            return tuple(typical), "typical", {}

        primary = station_ids[0]
        hook_km: dict[int, float] = {}
        hooked: list[LibraryRoute] = []
        for route in routes:
            hook = hooks.get((route.route_id, primary))
            if hook is not None:
                hooked.append(route)
                hook_km[route.route_id] = hook
        if hooked:
            return tuple(hooked), "hook", hook_km

    return tuple(routes), "all", {}


def _fits_corridor(
    route: LibraryRoute,
    *,
    fuel_start: float,
    q_today: float,
    norm: float,
    tank_volume: float,
) -> bool:
    """True when today's round-trip leaves the tank in ``[0, tank_volume]``."""
    fuel_end = fuel_start + q_today - burn_for_km(_daily_km(route), norm)
    return 0 <= fuel_end <= tank_volume


def _hook_choice_warnings(
    source: str,
    chosen: LibraryRoute,
    hook_km: Mapping[int, float],
    hook_threshold_km: float,
) -> tuple[str, ...]:
    if source != "hook":
        return ()
    hook = hook_km.get(chosen.route_id)
    if hook is None:
        return ()
    return ("hook_above_threshold",) if hook > hook_threshold_km else ()


def _choose_base_anchor(
    group: Sequence[LibraryRoute],
    *,
    source: str,
    hook_km: Mapping[int, float],
    hook_threshold_km: float,
    q_today: float,
) -> tuple[LibraryRoute, tuple[str, ...]]:
    """Base pick from a non-empty group: wash → min km, hook → min hook, else max freq."""
    if q_today == 0:
        chosen = min(group, key=lambda r: (_daily_km(r), r.route_id))
    elif source == "hook":
        chosen = min(group, key=lambda r: (hook_km[r.route_id], r.route_id))
    else:
        chosen = sorted(group, key=lambda r: (-r.frequency, r.route_id))[0]
    return chosen, _hook_choice_warnings(source, chosen, hook_km, hook_threshold_km)


def _lookahead_km_needed(
    *,
    fuel_before: float,
    q_today: float,
    q_next: float,
    tank_volume: float,
    norm: float,
) -> float:
    """Daily km to burn today so next fill fits; 0 if burn_needed ≤ 0."""
    headroom_needed = tank_volume - q_next
    fuel_after = fuel_before + q_today
    burn_needed = fuel_after - headroom_needed
    if burn_needed <= 0:
        return 0.0
    return burn_needed / norm * 100


def _tentative_fuel_end(
    fuel_before: float, q_today: float, route: LibraryRoute, norm: float
) -> float:
    return round(fuel_before + q_today - burn_for_km(_daily_km(route), norm), 2)


def _burn_in_reaches_headroom(
    *,
    fuel: float,
    free_days: Sequence[date],
    burn_routes: Sequence[LibraryRoute],
    target_fuel_max: float,
    tank_volume_liters: float,
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    at_date: date,
    fuel_to_issue: float,
    own_vehicle_id: int = 0,
) -> bool:
    if not free_days:
        return False
    _planned, _fuel, unsolvable = _plan_burn_in(
        fuel=fuel,
        free_days=free_days,
        burn_routes=burn_routes,
        target_fuel_max=target_fuel_max,
        tank_volume_liters=tank_volume_liters,
        norm_summer=norm_summer,
        norm_winter=norm_winter,
        season_switches=season_switches,
        at_date=at_date,
        fuel_to_issue=fuel_to_issue,
        own_vehicle_id=own_vehicle_id,
    )
    return unsolvable is None


def _first_known_point(
    station_ids: Sequence[int],
    station_coords: Mapping[int, GeoPoint],
) -> GeoPoint | None:
    for sid in station_ids:
        point = station_coords.get(sid)
        if point is not None:
            return point
    return None


def _target_bearing(
    today_station_ids: Sequence[int],
    next_station_ids: Sequence[int],
    station_coords: Mapping[int, GeoPoint] | None,
) -> float | None:
    if not station_coords:
        return None
    today = _first_known_point(today_station_ids, station_coords)
    nxt = _first_known_point(next_station_ids, station_coords)
    if today is None or nxt is None or today == nxt:
        return None
    return bearing_deg(today, nxt)


def _route_bearing(route: LibraryRoute) -> float | None:
    if route.point_a is None or route.point_b is None:
        return None
    return bearing_deg(route.point_a, route.point_b)


def _station_on_route(route: LibraryRoute, station_ids: Sequence[int]) -> bool:
    return any(sid in route.typical_station_ids for sid in station_ids)


def _direction_priority(
    route: LibraryRoute,
    *,
    today_station_ids: Sequence[int],
    target_bearing: float | None,
) -> int:
    """1 = typical + aligned, 2 = typical any heading, 3 = sufficient km / hook."""
    on_route = _station_on_route(route, today_station_ids)
    if on_route and target_bearing is not None:
        heading = _route_bearing(route)
        if heading is not None and angle_diff_deg(heading, target_bearing) <= _ALIGNED_MAX_ANGLE_DEG:
            return 1
    if on_route:
        return 2
    return 3


def _pick_min_sufficient(
    candidates: Sequence[LibraryRoute],
    *,
    km_needed: float,
    seed: int,
    today_station_ids: Sequence[int] = (),
    target_bearing: float | None = None,
) -> LibraryRoute | None:
    sufficient = [r for r in candidates if _daily_km(r) >= km_needed]
    if not sufficient:
        return None
    rng = Random(seed)
    tie = {r.route_id: rng.random() for r in sufficient}
    sufficient.sort(
        key=lambda r: (
            _direction_priority(
                r,
                today_station_ids=today_station_ids,
                target_bearing=target_bearing,
            ),
            _daily_km(r),
            -r.frequency,
            r.route_id,
            tie[r.route_id],
        )
    )
    return sufficient[0]


def _required_lookahead_km(
    *,
    chosen: LibraryRoute,
    fuel_before: float,
    q_today: float,
    next_fuel: _NextFuel | None,
    tank_volume_liters: float,
    norm: float,
    free_until_next: Sequence[date],
    burn_routes: Sequence[LibraryRoute],
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    own_vehicle_id: int = 0,
) -> float:
    """km_needed if today must burn extra for Q_next; 0 if normal pick is enough."""
    if next_fuel is None:
        return 0.0
    next_date, q_next, _next_stations = next_fuel
    km_needed = _lookahead_km_needed(
        fuel_before=fuel_before,
        q_today=q_today,
        q_next=q_next,
        tank_volume=tank_volume_liters,
        norm=norm,
    )
    if km_needed <= 0:
        return 0.0
    tentative_end = _tentative_fuel_end(fuel_before, q_today, chosen, norm)
    if _burn_in_reaches_headroom(
        fuel=tentative_end,
        free_days=free_until_next,
        burn_routes=burn_routes,
        target_fuel_max=tank_volume_liters - q_next,
        tank_volume_liters=tank_volume_liters,
        norm_summer=norm_summer,
        norm_winter=norm_winter,
        season_switches=season_switches,
        at_date=next_date,
        fuel_to_issue=q_next,
        own_vehicle_id=own_vehicle_id,
    ):
        return 0.0
    return km_needed


def _non_home_city_key(route: LibraryRoute) -> str:
    a_home = _is_home_base(route.addr_a)
    b_home = _is_home_base(route.addr_b)
    if a_home and not b_home:
        return _city_key(route.addr_b)
    if b_home and not a_home:
        return _city_key(route.addr_a)
    return _city_key(route.addr_b)


def _city_hint_key(
    routes: Sequence[LibraryRoute],
    *,
    own_vehicle_id: int,
    km_needed: float,
) -> str:
    own = [
        r
        for r in routes
        if r.vehicle_id == own_vehicle_id
        and (_is_home_base(r.addr_a) or _is_home_base(r.addr_b))
    ]
    if not own:
        return ""
    if km_needed > 0:
        best = min(own, key=lambda r: (abs(_daily_km(r) - km_needed), r.route_id))
    else:
        best = max(own, key=lambda r: (_daily_km(r), -r.route_id))
    return _non_home_city_key(best)


def _rank_key(
    route: LibraryRoute,
    *,
    own_vehicle_id: int,
    preferred_ids: frozenset[int],
    city_hint: str,
    invert_km: bool,
    today_station_ids: Sequence[int],
    target_bearing: float | None,
) -> tuple[int, int, int, int, int, int, int]:
    km = _daily_km(route)
    same_city = (not city_hint) or _non_home_city_key(route) == city_hint
    return (
        0 if route.vehicle_id == own_vehicle_id else 1,
        0 if route.route_id in preferred_ids else 1,
        0 if same_city else 1,
        _direction_priority(
            route,
            today_station_ids=today_station_ids,
            target_bearing=target_bearing,
        ),
        -km if invert_km else km,
        -route.frequency,
        route.route_id,
    )


def _rank_routes(
    candidates: Sequence[LibraryRoute],
    *,
    own_vehicle_id: int,
    preferred_ids: frozenset[int],
    city_hint: str,
    invert_km: bool = False,
    today_station_ids: Sequence[int] = (),
    target_bearing: float | None = None,
) -> LibraryRoute:
    return min(
        candidates,
        key=lambda r: _rank_key(
            r,
            own_vehicle_id=own_vehicle_id,
            preferred_ids=preferred_ids,
            city_hint=city_hint,
            invert_km=invert_km,
            today_station_ids=today_station_ids,
            target_bearing=target_bearing,
        ),
    )


def _select_anchor_route_lookahead(
    *,
    station_ids: Sequence[int],
    routes: Sequence[LibraryRoute],
    hooks: Mapping[tuple[int, int], float],
    hook_threshold_km: float,
    fuel_before: float,
    q_today: float,
    next_fuel: _NextFuel | None,
    tank_volume_liters: float,
    norm: float,
    free_until_next: Sequence[date],
    burn_routes: Sequence[LibraryRoute],
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    seed: int,
    station_coords: Mapping[int, GeoPoint] | None = None,
    own_vehicle_id: int = 0,
) -> tuple[LibraryRoute, tuple[str, ...]]:
    kwargs = dict(
        station_ids=station_ids,
        routes=routes,
        hooks=hooks,
        hook_threshold_km=hook_threshold_km,
        fuel_before=fuel_before,
        q_today=q_today,
        next_fuel=next_fuel,
        tank_volume_liters=tank_volume_liters,
        norm=norm,
        free_until_next=free_until_next,
        burn_routes=burn_routes,
        norm_summer=norm_summer,
        norm_winter=norm_winter,
        season_switches=season_switches,
        seed=seed,
        station_coords=station_coords,
        own_vehicle_id=own_vehicle_id,
    )
    if _home_base_routes(routes):
        return _select_anchor_route_fleet(**kwargs)
    return _select_anchor_route_legacy(**kwargs)


def _select_anchor_route_legacy(
    *,
    station_ids: Sequence[int],
    routes: Sequence[LibraryRoute],
    hooks: Mapping[tuple[int, int], float],
    hook_threshold_km: float,
    fuel_before: float,
    q_today: float,
    next_fuel: _NextFuel | None,
    tank_volume_liters: float,
    norm: float,
    free_until_next: Sequence[date],
    burn_routes: Sequence[LibraryRoute],
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    seed: int,
    station_coords: Mapping[int, GeoPoint] | None = None,
    own_vehicle_id: int = 0,
) -> tuple[LibraryRoute, tuple[str, ...]]:
    """Pre-fleet cascade for fixtures without a Kuznetskaya home base."""
    group, source, hook_km = _anchor_route_group(
        station_ids=station_ids,
        routes=routes,
        hooks=hooks,
    )
    in_corridor = tuple(
        r
        for r in group
        if _fits_corridor(
            r,
            fuel_start=fuel_before,
            q_today=q_today,
            norm=norm,
            tank_volume=tank_volume_liters,
        )
    )
    if in_corridor:
        chosen, warnings = _choose_base_anchor(
            in_corridor,
            source=source,
            hook_km=hook_km,
            hook_threshold_km=hook_threshold_km,
            q_today=q_today,
        )
    else:
        chosen = min(group, key=lambda r: (_daily_km(r), r.route_id))
        warnings = _hook_choice_warnings(source, chosen, hook_km, hook_threshold_km)
    km_needed = _required_lookahead_km(
        chosen=chosen,
        fuel_before=fuel_before,
        q_today=q_today,
        next_fuel=next_fuel,
        tank_volume_liters=tank_volume_liters,
        norm=norm,
        free_until_next=free_until_next,
        burn_routes=burn_routes,
        norm_summer=norm_summer,
        norm_winter=norm_winter,
        season_switches=season_switches,
        own_vehicle_id=own_vehicle_id,
    )
    if km_needed <= 0:
        return chosen, warnings
    next_station_ids = next_fuel[2] if next_fuel is not None else ()
    target_bearing = _target_bearing(station_ids, next_station_ids, station_coords)
    elongated = (
        _pick_min_sufficient(
            in_corridor,
            km_needed=km_needed,
            seed=seed,
            today_station_ids=station_ids,
            target_bearing=target_bearing,
        )
        if in_corridor
        else None
    )
    from_full_library = False
    if elongated is None:
        elongated = _pick_min_sufficient(
            group,
            km_needed=km_needed,
            seed=seed,
            today_station_ids=station_ids,
            target_bearing=target_bearing,
        )
    if elongated is None:
        elongated = _pick_min_sufficient(
            routes,
            km_needed=km_needed,
            seed=seed,
            today_station_ids=station_ids,
            target_bearing=target_bearing,
        )
        from_full_library = elongated is not None
    if elongated is None:
        return chosen, warnings
    extra = (
        ("balance_route",)
        if from_full_library or elongated.route_id != chosen.route_id
        else ()
    )
    return elongated, warnings + extra


def _select_anchor_route_fleet(
    *,
    station_ids: Sequence[int],
    routes: Sequence[LibraryRoute],
    hooks: Mapping[tuple[int, int], float],
    hook_threshold_km: float,
    fuel_before: float,
    q_today: float,
    next_fuel: _NextFuel | None,
    tank_volume_liters: float,
    norm: float,
    free_until_next: Sequence[date],
    burn_routes: Sequence[LibraryRoute],
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    seed: int,
    station_coords: Mapping[int, GeoPoint] | None = None,
    own_vehicle_id: int = 0,
) -> tuple[LibraryRoute, tuple[str, ...]]:
    del seed
    fleet = _fleet_pool(routes)
    group, source, hook_km = _anchor_route_group(
        station_ids=station_ids,
        routes=routes,
        hooks=hooks,
    )
    preferred_ids = frozenset(r.route_id for r in group)
    in_corridor = tuple(
        r
        for r in fleet
        if _fits_corridor(
            r,
            fuel_start=fuel_before,
            q_today=q_today,
            norm=norm,
            tank_volume=tank_volume_liters,
        )
    )
    next_station_ids = next_fuel[2] if next_fuel is not None else ()
    target_bearing = _target_bearing(station_ids, next_station_ids, station_coords)

    km_needed = 0.0
    if next_fuel is not None and in_corridor:
        tentative = _rank_routes(
            in_corridor,
            own_vehicle_id=own_vehicle_id,
            preferred_ids=preferred_ids,
            city_hint=_city_hint_key(
                in_corridor, own_vehicle_id=own_vehicle_id, km_needed=0.0
            ),
            today_station_ids=station_ids,
            target_bearing=target_bearing,
        )
        km_needed = _required_lookahead_km(
            chosen=tentative,
            fuel_before=fuel_before,
            q_today=q_today,
            next_fuel=next_fuel,
            tank_volume_liters=tank_volume_liters,
            norm=norm,
            free_until_next=free_until_next,
            burn_routes=burn_routes,
            norm_summer=norm_summer,
            norm_winter=norm_winter,
            season_switches=season_switches,
            own_vehicle_id=own_vehicle_id,
        )
    elif next_fuel is not None:
        km_needed = _lookahead_km_needed(
            fuel_before=fuel_before,
            q_today=q_today,
            q_next=next_fuel[1],
            tank_volume=tank_volume_liters,
            norm=norm,
        )

    city_hint = _city_hint_key(
        fleet, own_vehicle_id=own_vehicle_id, km_needed=km_needed
    )
    sufficient = (
        tuple(r for r in in_corridor if _daily_km(r) >= km_needed)
        if km_needed > 0
        else in_corridor
    )

    invert_km = False
    manual = False
    if sufficient:
        pool: Sequence[LibraryRoute] = sufficient
    elif in_corridor:
        pool = in_corridor
        invert_km = True
        manual = True
    else:
        pool = fleet or group
        manual = True

    chosen = _rank_routes(
        pool,
        own_vehicle_id=own_vehicle_id,
        preferred_ids=preferred_ids,
        city_hint=city_hint,
        invert_km=invert_km,
        today_station_ids=station_ids,
        target_bearing=target_bearing,
    )

    warnings: list[str] = list(
        _hook_choice_warnings(source, chosen, hook_km, hook_threshold_km)
    )
    if chosen.vehicle_id != own_vehicle_id:
        warnings.append("borrowed_route")
    elif chosen.route_id not in preferred_ids:
        warnings.append("balance_route")
    elif km_needed > 0:
        preferred_in = tuple(r for r in in_corridor if r.route_id in preferred_ids)
        if preferred_in:
            base, _ = _choose_base_anchor(
                preferred_in,
                source=source,
                hook_km=hook_km,
                hook_threshold_km=hook_threshold_km,
                q_today=q_today,
            )
            if chosen.route_id != base.route_id:
                warnings.append("balance_route")
    if manual:
        warnings.append(_MANUAL_REASON)
    return chosen, tuple(warnings)


def _plan_burn_in(
    *,
    fuel: float,
    free_days: Sequence[date],
    burn_routes: Sequence[LibraryRoute],
    target_fuel_max: float,
    tank_volume_liters: float,
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    at_date: date,
    fuel_to_issue: float,
    own_vehicle_id: int = 0,
) -> tuple[list[tuple[date, LibraryRoute]], float, UnsolvableInfo | None]:
    fuel_cur = fuel
    if fuel_cur <= target_fuel_max + 1e-9:
        return [], fuel_cur, None

    if not burn_routes or not free_days:
        return [], fuel_cur, UnsolvableInfo(
            reason="insufficient_headroom",
            at_date=at_date,
            fuel_before=fuel_cur,
            fuel_to_issue=fuel_to_issue,
            tank_volume=tank_volume_liters,
            free_weekdays=len(free_days),
            detail=(
                f"cannot reach headroom ≤ {target_fuel_max} before "
                f"{at_date.isoformat()} (free_weekdays={len(free_days)})"
            ),
        )

    planned: list[tuple[date, LibraryRoute]] = []
    for day in free_days:
        if fuel_cur <= target_fuel_max + 1e-9:
            break
        norm = norm_for(
            day,
            norm_summer=norm_summer,
            norm_winter=norm_winter,
            switches=season_switches,
        )
        candidates: list[tuple[LibraryRoute, float, float]] = []
        for candidate in burn_routes:
            burn = burn_for_km(_daily_km(candidate), norm)
            nxt = round(fuel_cur - burn, 2)
            if 0.0 <= nxt <= tank_volume_liters:
                candidates.append((candidate, burn, nxt))
        if not candidates:
            continue

        reaching = [c for c in candidates if c[2] <= target_fuel_max + 1e-9]
        if reaching:
            # Land in headroom: min sufficient daily km; own beats foreign.
            reaching.sort(
                key=lambda c: (
                    0 if c[0].vehicle_id == own_vehicle_id else 1,
                    _daily_km(c[0]),
                    -c[0].frequency,
                    c[0].route_id,
                )
            )
            chosen, _burn, nxt = reaching[0]
        else:
            # Need more days: maximize safe burn; own beats foreign at same burn.
            candidates.sort(
                key=lambda c: (
                    -c[1],
                    0 if c[0].vehicle_id == own_vehicle_id else 1,
                    -c[0].frequency,
                    c[0].route_id,
                )
            )
            chosen, _burn, nxt = candidates[0]
        planned.append((day, chosen))
        fuel_cur = nxt

    if fuel_cur > target_fuel_max + 1e-9:
        return planned, fuel_cur, UnsolvableInfo(
            reason="insufficient_headroom",
            at_date=at_date,
            fuel_before=fuel_cur,
            fuel_to_issue=fuel_to_issue,
            tank_volume=tank_volume_liters,
            free_weekdays=len(free_days),
            detail=(
                f"after {len(planned)} burn-in days fuel={fuel_cur} still "
                f"> {target_fuel_max} before {at_date.isoformat()}"
            ),
        )
    return planned, fuel_cur, None


def _emit_maybe_manual(
    *,
    day: date,
    route: LibraryRoute,
    fuel_start: float,
    fuel_issued: float,
    odometer_start: int,
    driver_id: int,
    tank_volume_liters: float,
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    day_warnings: tuple[str, ...],
    detail: str,
    force_manual: bool = False,
    own_vehicle_id: int = 0,
    catalog: Sequence[LibraryRoute] = (),
) -> tuple[WaybillDay, float, int, ProblematicDay | None]:
    """Emit a day; on corridor/headroom failure keep the day as manual_intervention."""
    warnings = day_warnings
    if force_manual and _MANUAL_REASON not in warnings:
        warnings = warnings + (_MANUAL_REASON,)

    def emit(allow_breach: bool, warn: tuple[str, ...]) -> tuple[WaybillDay | None, float, int]:
        return _emit_day(
            day=day,
            route=route,
            fuel_start=fuel_start,
            fuel_issued=fuel_issued,
            odometer_start=odometer_start,
            driver_id=driver_id,
            tank_volume_liters=tank_volume_liters,
            norm_summer=norm_summer,
            norm_winter=norm_winter,
            season_switches=season_switches,
            day_warnings=warn,
            allow_corridor_breach=allow_breach,
            own_vehicle_id=own_vehicle_id,
            catalog=catalog,
        )

    wb, fuel, odometer = emit(force_manual, warnings)
    first_failed = wb is None
    if first_failed:
        if _MANUAL_REASON not in warnings:
            warnings = warnings + (_MANUAL_REASON,)
        wb, fuel, odometer = emit(True, warnings)
    if wb is None:
        raise RuntimeError(f"manual emit must succeed for {day.isoformat()}")
    if not (force_manual or first_failed):
        return wb, fuel, odometer, None
    problem = ProblematicDay(
        date=day,
        reason=_MANUAL_REASON,
        detail=detail,
        fuel_before=fuel_start,
        fuel_to_issue=fuel_issued,
        tank_volume=tank_volume_liters,
    )
    return wb, fuel, odometer, problem


def _emit_day(
    *,
    day: date,
    route: LibraryRoute,
    fuel_start: float,
    fuel_issued: float,
    odometer_start: int,
    driver_id: int,
    tank_volume_liters: float,
    norm_summer: float,
    norm_winter: float,
    season_switches: Sequence[SeasonSwitch],
    day_warnings: tuple[str, ...],
    allow_corridor_breach: bool = False,
    own_vehicle_id: int = 0,
    catalog: Sequence[LibraryRoute] = (),
) -> tuple[WaybillDay | None, float, int]:
    oriented = _orient_home_round_trip(
        route, catalog=catalog, own_vehicle_id=own_vehicle_id
    )
    norm = norm_for(
        day,
        norm_summer=norm_summer,
        norm_winter=norm_winter,
        switches=season_switches,
    )
    km = _daily_km(route)
    try:
        tank: TankState = apply_day(
            day,
            fuel_start=fuel_start,
            fuel_issued=fuel_issued,
            km=km,
            odometer_start=odometer_start,
            norm_per_100km=norm,
            tank_volume_liters=tank_volume_liters,
        )
    except BalanceViolation:
        if not allow_corridor_breach:
            return None, fuel_start, odometer_start
        fuel_end = round(fuel_start + fuel_issued - burn_for_km(km, norm), 2)
        tank = TankState(
            date=day,
            fuel_start=fuel_start,
            fuel_issued=fuel_issued,
            fuel_end=fuel_end,
            km=km,
            odometer_start=odometer_start,
            odometer_end=odometer_start + km,
        )

    waybill = WaybillDay(
        date=day,
        driver_id=driver_id,
        route=_to_route_ref(oriented, own_vehicle_id=own_vehicle_id),
        tank=tank,
        source="auto",
        warnings=day_warnings,
        legs=_round_trip_legs(oriented, own_vehicle_id=own_vehicle_id),
    )
    return waybill, tank.fuel_end, tank.odometer_end


def _dedupe_warnings(codes: Sequence[str]) -> tuple[str, ...]:
    seen: list[str] = []
    for code in codes:
        if code not in seen:
            seen.append(code)
    return tuple(seen)
