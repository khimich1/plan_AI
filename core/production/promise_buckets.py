"""Weekly promise-bucket math: tracks, weeks, allocation (no I/O, no app.*)."""

from __future__ import annotations

import math
from collections.abc import Callable, Collection, Mapping, Sequence
from dataclasses import dataclass
from datetime import date, timedelta

from core.production.capacity import TRACKS_PER_DAY_HARD_CAP
from core.production_capacity import MAX_TRACK_LENGTH_M
from core.work_calendar import is_working_day

# Locked assumption 2 (Task 0 / human): coefficient on m/101, 0% markup.
DEFAULT_TRACK_BUFFER = 1.0
DEFAULT_PROMISE_TRACKS_PER_DAY = 3
_WEEK_DAYS = 7
_HORIZON_SCAN_DAYS = 366 * 2

OccupancyMap = Mapping[date | str, int]


class OccupancyUnavailableError(ValueError):
    """Plan occupancy is missing — fail closed, never treat as 'all free'."""


@dataclass(frozen=True, slots=True)
class WeekBucket:
    week_start: date
    workdays: int
    capacity: int
    planned: int
    promised: int
    held: int

    @property
    def free(self) -> int:
        return max(0, self.capacity - self.planned - self.promised)


@dataclass(frozen=True, slots=True)
class PromiseWindow:
    from_week: date
    to_week: date
    promised_date: date
    allocations: tuple[tuple[date, int], ...]


@dataclass(frozen=True, slots=True)
class PourPlan:
    first_pour_date: date
    first_pour_free: int
    solo_date: date
    solo_week_end_date: date
    allocations: tuple[tuple[date, int], ...]  # days with take > 0


@dataclass(frozen=True, slots=True)
class PromiseQuote:
    tracks: int
    solo_days: int
    solo_date: date | None
    solo_week_end_date: date | None
    earliest_start_week: date | None
    first_pour_date: date | None
    first_pour_free: int
    window: PromiseWindow | None
    weeks: tuple[WeekBucket, ...]
    knob: int
    buffer: float


def estimate_tracks(
    total_length_m: float,
    buffer: float = DEFAULT_TRACK_BUFFER,
) -> int:
    """tracks = ceil(Σ(length_m × qty) / MAX_TRACK_LENGTH_M × buffer)."""
    if total_length_m <= 0 or buffer <= 0:
        return 0
    raw = (float(total_length_m) / MAX_TRACK_LENGTH_M) * float(buffer)
    return max(0, int(math.ceil(raw - 1e-12)))


def clamp_promise_knob(value: int) -> int:
    return max(1, min(int(value), TRACKS_PER_DAY_HARD_CAP))


def day_free(occupied: int, knob: int) -> int:
    return max(0, int(knob) - max(0, int(occupied)))


def solo_days(tracks: int, knob: int) -> int:
    if tracks <= 0:
        return 0
    return math.ceil(tracks / clamp_promise_knob(knob))


def iso_week_start(day: date) -> date:
    return day - timedelta(days=day.weekday())


def workday_predicate(
    holidays: Collection[date] | None = None,
    extra_workdays: Collection[date] | None = None,
) -> Callable[[date], bool]:
    """Bind work_calendar.is_working_day to explicit sets (no file I/O)."""
    hol = set(holidays or ())
    extra = set(extra_workdays or ())
    return lambda day: is_working_day(day, hol, extra)


def _default_is_workday(day: date) -> bool:
    return day.weekday() < 5


def _resolve_workday(
    is_workday: Callable[[date], bool] | None,
) -> Callable[[date], bool]:
    return is_workday if is_workday is not None else _default_is_workday


def last_workday_of_week(
    week_start: date,
    is_workday: Callable[[date], bool],
) -> date | None:
    last: date | None = None
    for offset in range(_WEEK_DAYS):
        day = week_start + timedelta(days=offset)
        if is_workday(day):
            last = day
    return last


def nth_workday_on_or_after(
    start: date,
    n: int,
    is_workday: Callable[[date], bool],
) -> date | None:
    if n < 1:
        return None
    found = 0
    current = start
    for _ in range(_HORIZON_SCAN_DAYS):
        if is_workday(current):
            found += 1
            if found == n:
                return current
        current += timedelta(days=1)
    return None


def _require_occupancy(occupancy: OccupancyMap | None) -> OccupancyMap:
    if occupancy is None or not isinstance(occupancy, Mapping):
        raise OccupancyUnavailableError(
            "Недоступна занятость плана — котировка остановлена (fail-closed)."
        )
    return occupancy


def _lookup_int(mapping: OccupancyMap, day: date) -> int:
    if day in mapping:
        return int(mapping[day])
    iso = day.isoformat()
    if iso in mapping:
        return int(mapping[iso])
    return 0


def _count_remaining_workdays(
    range_start: date,
    range_end: date,
    occupancy: OccupancyMap,
    is_workday: Callable[[date], bool],
) -> tuple[int, int]:
    workdays = 0
    planned = 0
    day = range_start
    while day <= range_end:
        if is_workday(day):
            workdays += 1
            planned += _lookup_int(occupancy, day)
        day += timedelta(days=1)
    return workdays, planned


def build_weeks(
    today: date,
    occupancy: OccupancyMap | None,
    *,
    promised_by_week: OccupancyMap | None = None,
    held_by_week: OccupancyMap | None = None,
    knob: int = DEFAULT_PROMISE_TRACKS_PER_DAY,
    week_count: int = 12,
    is_workday: Callable[[date], bool] | None = None,
) -> list[WeekBucket]:
    """ISO weeks from tomorrow (first week is partial). Occupancy is required."""
    occ = _require_occupancy(occupancy)
    workday_fn = _resolve_workday(is_workday)
    k = clamp_promise_knob(knob)
    promised = promised_by_week or {}
    held = held_by_week or {}
    tomorrow = today + timedelta(days=1)
    cursor = iso_week_start(tomorrow)
    weeks: list[WeekBucket] = []
    for _ in range(max(0, int(week_count))):
        sunday = cursor + timedelta(days=_WEEK_DAYS - 1)
        workdays, planned = _count_remaining_workdays(
            max(tomorrow, cursor), sunday, occ, workday_fn
        )
        weeks.append(
            WeekBucket(
                week_start=cursor,
                workdays=workdays,
                capacity=workdays * k,
                planned=planned,
                promised=_lookup_int(promised, cursor),
                held=_lookup_int(held, cursor),
            )
        )
        cursor += timedelta(days=_WEEK_DAYS)
    return weeks


def _whole_window(
    week: WeekBucket,
    tracks: int,
    is_workday: Callable[[date], bool],
) -> PromiseWindow | None:
    promised_date = last_workday_of_week(week.week_start, is_workday)
    if promised_date is None:
        return None
    return PromiseWindow(
        from_week=week.week_start,
        to_week=week.week_start,
        promised_date=promised_date,
        allocations=((week.week_start, tracks),),
    )


def _greedy_window(
    weeks: Sequence[WeekBucket],
    tracks: int,
    is_workday: Callable[[date], bool],
) -> PromiseWindow | None:
    start = next((i for i, week in enumerate(weeks) if week.free > 0), None)
    if start is None:
        return None
    remaining = tracks
    allocs: list[tuple[date, int]] = []
    last_week: WeekBucket | None = None
    for week in weeks[start:]:
        take = min(week.free, remaining)
        if take > 0:
            allocs.append((week.week_start, take))
            last_week = week
            remaining -= take
        if remaining == 0:
            break
    if remaining > 0 or last_week is None:
        return None
    promised_date = last_workday_of_week(last_week.week_start, is_workday)
    if promised_date is None:
        return None
    return PromiseWindow(
        from_week=allocs[0][0],
        to_week=last_week.week_start,
        promised_date=promised_date,
        allocations=tuple(allocs),
    )


def allocate(
    tracks: int,
    weeks: Sequence[WeekBucket],
    *,
    is_workday: Callable[[date], bool] | None = None,
) -> PromiseWindow | None:
    """Whole into first week with free >= tracks; else greedy window if too big."""
    if tracks <= 0 or not weeks:
        return None
    workday_fn = _resolve_workday(is_workday)
    max_capacity = max(week.capacity for week in weeks)
    if tracks <= max_capacity:
        for week in weeks:
            if week.free >= tracks:
                return _whole_window(week, tracks, workday_fn)
        return None
    return _greedy_window(weeks, tracks, workday_fn)


def week_allocations(pour: PourPlan) -> tuple[tuple[date, int], ...]:
    """Sum daily take by ISO Monday for the hold/promise journal."""
    totals: dict[date, int] = {}
    order: list[date] = []
    for day, take in pour.allocations:
        monday = iso_week_start(day)
        if monday not in totals:
            order.append(monday)
            totals[monday] = 0
        totals[monday] += take
    return tuple((monday, totals[monday]) for monday in order)


def pack_pour(
    tracks: int,
    occupancy: OccupancyMap,
    weeks: Sequence[WeekBucket],
    *,
    today: date,
    knob: int,
    is_workday: Callable[[date], bool],
    horizon_days: int = _HORIZON_SCAN_DAYS,
) -> PourPlan | None:
    """Daily pour from tomorrow: day remainder, week.free leftover, remaining tracks."""
    if tracks <= 0:
        return None
    k = clamp_promise_knob(knob)
    buckets = {week.week_start: week for week in weeks}
    taken_in_week: dict[date, int] = {}
    remaining = int(tracks)
    allocs: list[tuple[date, int]] = []
    for offset in range(1, max(1, int(horizon_days)) + 1):
        day = today + timedelta(days=offset)
        if not is_workday(day):
            continue
        day_left = day_free(_lookup_int(occupancy, day), k)
        week = iso_week_start(day)
        bucket = buckets.get(week)
        week_left = (bucket.free - taken_in_week.get(week, 0)) if bucket else 0
        take = min(remaining, day_left, week_left)
        if take > 0:
            allocs.append((day, take))
            remaining -= take
            taken_in_week[week] = taken_in_week.get(week, 0) + take
        if remaining == 0:
            break
    if not allocs or remaining > 0:
        return None
    first = allocs[0][0]
    solo = allocs[-1][0]
    week_end = last_workday_of_week(iso_week_start(solo), is_workday)
    if week_end is None:
        return None
    return PourPlan(
        first_pour_date=first,
        first_pour_free=day_free(_lookup_int(occupancy, first), k),
        solo_date=solo,
        solo_week_end_date=week_end,
        allocations=tuple(allocs),
    )


def _window_from_pour(pour: PourPlan) -> PromiseWindow:
    mondays = week_allocations(pour)
    return PromiseWindow(
        from_week=mondays[0][0],
        to_week=mondays[-1][0],
        promised_date=pour.solo_week_end_date,
        allocations=mondays,
    )


def build_quote(
    total_length_m: float,
    weeks: Sequence[WeekBucket],
    *,
    today: date,
    knob: int = DEFAULT_PROMISE_TRACKS_PER_DAY,
    buffer: float = DEFAULT_TRACK_BUFFER,
    is_workday: Callable[[date], bool] | None = None,
    occupancy: OccupancyMap | None = None,
) -> PromiseQuote:
    workday_fn = _resolve_workday(is_workday)
    k = clamp_promise_knob(knob)
    tracks = estimate_tracks(total_length_m, buffer=buffer)
    days = solo_days(tracks, k)
    occ: OccupancyMap = occupancy if occupancy is not None else {}
    pour = pack_pour(
        tracks, occ, weeks, today=today, knob=k, is_workday=workday_fn
    )
    if pour is None:
        finish = nth_workday_on_or_after(today + timedelta(days=1), days, workday_fn)
        week_end = (
            last_workday_of_week(iso_week_start(finish), workday_fn) if finish else None
        )
        return PromiseQuote(
            tracks=tracks,
            solo_days=days,
            solo_date=finish,
            solo_week_end_date=week_end,
            earliest_start_week=None,
            first_pour_date=None,
            first_pour_free=0,
            window=None,
            weeks=tuple(weeks),
            knob=k,
            buffer=buffer,
        )
    return PromiseQuote(
        tracks=tracks,
        solo_days=days,
        solo_date=pour.solo_date,
        solo_week_end_date=pour.solo_week_end_date,
        earliest_start_week=iso_week_start(pour.first_pour_date),
        first_pour_date=pour.first_pour_date,
        first_pour_free=pour.first_pour_free,
        window=_window_from_pour(pour),
        weeks=tuple(weeks),
        knob=k,
        buffer=buffer,
    )


__all__ = [
    "DEFAULT_PROMISE_TRACKS_PER_DAY",
    "DEFAULT_TRACK_BUFFER",
    "OccupancyUnavailableError",
    "PourPlan",
    "PromiseQuote",
    "PromiseWindow",
    "WeekBucket",
    "allocate",
    "build_quote",
    "build_weeks",
    "clamp_promise_knob",
    "day_free",
    "estimate_tracks",
    "iso_week_start",
    "last_workday_of_week",
    "nth_workday_on_or_after",
    "pack_pour",
    "solo_days",
    "week_allocations",
    "workday_predicate",
]
