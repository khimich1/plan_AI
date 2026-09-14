"""Pure-domain tests for weekly promise buckets (no I/O, no app.*)."""

from __future__ import annotations

import ast
import inspect
from datetime import date
from pathlib import Path

import pytest

from core.production.promise_buckets import (
    DEFAULT_PROMISE_TRACKS_PER_DAY,
    DEFAULT_TRACK_BUFFER,
    OccupancyUnavailableError,
    PromiseWindow,
    WeekBucket,
    allocate,
    build_quote,
    build_weeks,
    day_free,
    estimate_tracks,
    iso_week_start,
    last_workday_of_week,
    pack_pour,
    solo_days,
    week_allocations,
    workday_predicate,
)
from core.production_capacity import MAX_TRACK_LENGTH_M

MODULE_PATH = (
    Path(__file__).resolve().parents[1] / "core" / "production" / "promise_buckets.py"
)

# Wednesday: remaining week is Thu–Sun (2 default workdays).
TODAY = date(2026, 9, 2)
WEEK_0 = date(2026, 8, 31)  # Mon of the partial week
WEEK_1 = date(2026, 9, 7)
WEEK_2 = date(2026, 9, 14)


def _bucket(
    week_start: date,
    *,
    workdays: int,
    planned: int = 0,
    promised: int = 0,
    held: int = 0,
    knob: int = DEFAULT_PROMISE_TRACKS_PER_DAY,
) -> WeekBucket:
    return WeekBucket(
        week_start=week_start,
        workdays=workdays,
        capacity=workdays * knob,
        planned=planned,
        promised=promised,
        held=held,
    )


def test_default_buffer_is_one() -> None:
    assert DEFAULT_TRACK_BUFFER == 1.0
    params = inspect.signature(estimate_tracks).parameters
    assert params["buffer"].default == 1.0


def test_estimate_tracks_default_buffer_is_not_115() -> None:
    # 101 м × 1.0 → 1 дорожка; при buffer=1.15 было бы 2.
    assert estimate_tracks(MAX_TRACK_LENGTH_M) == 1
    assert estimate_tracks(MAX_TRACK_LENGTH_M, buffer=1.15) == 2


def test_estimate_tracks_buffer_is_parameter() -> None:
    length = 2 * MAX_TRACK_LENGTH_M
    assert estimate_tracks(length, buffer=DEFAULT_TRACK_BUFFER) == 2
    assert estimate_tracks(length, buffer=1.0) == 2
    assert estimate_tracks(length, buffer=1.15) == 3


def test_estimate_tracks_zero_and_negative() -> None:
    assert estimate_tracks(0) == 0
    assert estimate_tracks(-10) == 0
    assert estimate_tracks(MAX_TRACK_LENGTH_M, buffer=0) == 0


def test_estimate_tracks_ceils_fraction() -> None:
    assert estimate_tracks(MAX_TRACK_LENGTH_M + 0.1, buffer=1.0) == 2


def test_week_bucket_free_floors_at_zero() -> None:
    week = _bucket(WEEK_1, workdays=5, planned=20)
    assert week.capacity == 15
    assert week.free == 0


def test_week_bucket_held_not_subtracted_from_free() -> None:
    week = _bucket(WEEK_1, workdays=5, planned=5, promised=5, held=10)
    assert week.free == 5


def test_allocate_whole_only_skips_fragmented_weeks() -> None:
    weeks = [
        _bucket(WEEK_0, workdays=5, planned=10),  # free 5
        _bucket(WEEK_1, workdays=5, planned=10),  # free 5
        _bucket(WEEK_2, workdays=5),  # free 15
    ]
    window = allocate(8, weeks)
    assert window is not None
    assert window.from_week == WEEK_2
    assert window.to_week == WEEK_2
    assert window.allocations == ((WEEK_2, 8),)
    assert window.promised_date == date(2026, 9, 18)


def test_allocate_large_kp_greedy_window() -> None:
    weeks = [_bucket(WEEK_1, workdays=5), _bucket(WEEK_2, workdays=5)]
    window = allocate(20, weeks)
    assert window is not None
    assert window.from_week == WEEK_1
    assert window.to_week == WEEK_2
    assert window.allocations == ((WEEK_1, 15), (WEEK_2, 5))
    assert window.promised_date == date(2026, 9, 18)


def test_allocate_returns_none_when_horizon_too_short() -> None:
    weeks = [_bucket(WEEK_1, workdays=5)]
    assert allocate(20, weeks) is None


def test_allocate_zero_tracks_returns_none() -> None:
    assert allocate(0, [_bucket(WEEK_1, workdays=5)]) is None


def test_promised_date_is_last_workday_of_last_week() -> None:
    friday_off = workday_predicate(holidays={date(2026, 9, 11)})
    weeks = [_bucket(WEEK_1, workdays=4)]
    window = allocate(4, weeks, is_workday=friday_off)
    assert window is not None
    assert window.promised_date == date(2026, 9, 10)
    assert last_workday_of_week(WEEK_1, friday_off) == date(2026, 9, 10)


def test_promised_date_uses_extra_workday_weekend() -> None:
    saturday_on = workday_predicate(extra_workdays={date(2026, 9, 12)})
    weeks = [_bucket(WEEK_1, workdays=6)]
    window = allocate(3, weeks, is_workday=saturday_on)
    assert window is not None
    assert window.promised_date == date(2026, 9, 12)


def test_build_weeks_partial_current_week() -> None:
    weeks = build_weeks(TODAY, occupancy={}, week_count=2)
    first = weeks[0]
    assert first.week_start == WEEK_0
    assert first.workdays == 2  # Thu 3 + Fri 4
    assert first.capacity == 6
    assert first.planned == 0
    assert first.free == 6
    assert weeks[1].week_start == WEEK_1
    assert weeks[1].workdays == 5
    assert weeks[1].capacity == 15


def test_build_weeks_holiday_week() -> None:
    friday_off = workday_predicate(holidays={date(2026, 9, 11)})
    weeks = build_weeks(
        date(2026, 9, 6),  # Sunday → first week is Mon 7–Sun 13
        occupancy={},
        week_count=1,
        is_workday=friday_off,
    )
    assert weeks[0].week_start == WEEK_1
    assert weeks[0].workdays == 4
    assert weeks[0].capacity == 12


def test_build_weeks_extra_workday_in_partial_week() -> None:
    saturday_on = workday_predicate(extra_workdays={date(2026, 9, 5)})
    weeks = build_weeks(TODAY, occupancy={}, week_count=1, is_workday=saturday_on)
    assert weeks[0].workdays == 3  # Thu, Fri, Sat
    assert weeks[0].capacity == 9


def test_build_weeks_planned_only_on_remaining_days() -> None:
    occupancy = {
        date(2026, 8, 31): 3,  # Mon — already past, must not count
        "2026-09-01": 3,  # Tue
        date(2026, 9, 2): 3,  # today
        date(2026, 9, 3): 2,  # tomorrow Thu
        "2026-09-04": 1,  # Fri
    }
    weeks = build_weeks(TODAY, occupancy, week_count=1)
    assert weeks[0].planned == 3
    assert weeks[0].capacity == 6
    assert weeks[0].free == 3


def test_build_weeks_promised_and_held() -> None:
    weeks = build_weeks(
        TODAY,
        occupancy={},
        promised_by_week={WEEK_0: 2, WEEK_1: 4},
        held_by_week={WEEK_0: 7},
        week_count=2,
    )
    assert weeks[0].promised == 2
    assert weeks[0].held == 7
    assert weeks[0].free == 4  # capacity 6 − promised 2; held ignored
    assert weeks[1].promised == 4
    assert weeks[1].held == 0
    assert weeks[1].free == 11


def test_build_weeks_occupancy_none_is_fail_closed() -> None:
    with pytest.raises(OccupancyUnavailableError, match="занятость"):
        build_weeks(TODAY, occupancy=None)


def test_build_weeks_occupancy_not_mapping_is_fail_closed() -> None:
    with pytest.raises(OccupancyUnavailableError):
        build_weeks(TODAY, occupancy="broken")  # type: ignore[arg-type]


def test_solo_days_from_tracks_and_knob() -> None:
    assert solo_days(1, 3) == 1
    assert solo_days(3, 3) == 1
    assert solo_days(4, 3) == 2
    assert solo_days(0, 3) == 0


def test_iso_week_start_is_monday() -> None:
    assert iso_week_start(date(2026, 9, 2)) == WEEK_0
    assert iso_week_start(WEEK_0) == WEEK_0
    assert iso_week_start(date(2026, 9, 6)) == WEEK_0


def test_build_quote_assembles_spec_fields() -> None:
    today = date(2026, 9, 6)  # Sunday → first basket is a full week
    weeks = build_weeks(today, occupancy={}, week_count=3)
    quote = build_quote(
        20 * MAX_TRACK_LENGTH_M,
        weeks,
        today=today,
        knob=DEFAULT_PROMISE_TRACKS_PER_DAY,
    )
    assert quote.tracks == 20
    assert quote.buffer == DEFAULT_TRACK_BUFFER
    assert quote.knob == 3
    assert quote.solo_days == 7
    assert quote.solo_date == date(2026, 9, 15)
    assert quote.solo_week_end_date == date(2026, 9, 18)
    assert quote.window is not None
    assert quote.earliest_start_week == WEEK_1
    assert quote.window.allocations == ((WEEK_1, 15), (WEEK_2, 5))
    assert quote.window.promised_date == date(2026, 9, 18)
    assert quote.weeks == tuple(weeks)


def test_build_quote_large_kp_starts_in_partial_week() -> None:
    weeks = build_weeks(TODAY, occupancy={}, week_count=3)
    quote = build_quote(
        20 * MAX_TRACK_LENGTH_M,
        weeks,
        today=TODAY,
        knob=DEFAULT_PROMISE_TRACKS_PER_DAY,
    )
    assert quote.window is not None
    assert quote.earliest_start_week == WEEK_0
    assert quote.window.allocations == ((WEEK_0, 6), (WEEK_1, 14))


EXAMPLE_A_TODAY = date(2026, 9, 4)
EXAMPLE_A_OCCUPANCY = {
    date(2026, 9, 7): 3,
    date(2026, 9, 8): 3,
    date(2026, 9, 9): 1,
}
WEEKDAYS = lambda day: day.weekday() < 5


def test_day_free_floors_at_zero() -> None:
    assert day_free(0, 3) == 3
    assert day_free(1, 3) == 2
    assert day_free(3, 3) == 0
    assert day_free(4, 3) == 0


def test_pack_pour_example_a() -> None:
    weeks = build_weeks(
        EXAMPLE_A_TODAY, EXAMPLE_A_OCCUPANCY, week_count=4, is_workday=WEEKDAYS
    )
    pour = pack_pour(
        5,
        EXAMPLE_A_OCCUPANCY,
        weeks,
        today=EXAMPLE_A_TODAY,
        knob=3,
        is_workday=WEEKDAYS,
    )
    assert pour is not None
    assert pour.first_pour_date == date(2026, 9, 9)
    assert pour.first_pour_free == 2
    assert pour.allocations == ((date(2026, 9, 9), 2), (date(2026, 9, 10), 3))
    assert pour.solo_date == date(2026, 9, 10)
    assert pour.solo_week_end_date == date(2026, 9, 11)
    assert week_allocations(pour) == ((WEEK_1, 5),)


def test_pack_pour_example_b_skips_week_eaten_by_promised() -> None:
    weeks = build_weeks(
        EXAMPLE_A_TODAY,
        EXAMPLE_A_OCCUPANCY,
        promised_by_week={WEEK_1: 8},
        week_count=4,
        is_workday=WEEKDAYS,
    )
    assert weeks[1].week_start == WEEK_1
    assert weeks[1].free == 0
    pour = pack_pour(
        5,
        EXAMPLE_A_OCCUPANCY,
        weeks,
        today=EXAMPLE_A_TODAY,
        knob=3,
        is_workday=WEEKDAYS,
    )
    assert pour is not None
    assert pour.first_pour_date >= WEEK_2
    assert all(day >= WEEK_2 for day, _take in pour.allocations)
    assert week_allocations(pour)[0][0] != WEEK_1


def test_pack_pour_horizon_exhausted_returns_none() -> None:
    weeks = build_weeks(EXAMPLE_A_TODAY, {}, week_count=1, is_workday=WEEKDAYS)
    assert (
        pack_pour(
            80,
            {},
            weeks,
            today=EXAMPLE_A_TODAY,
            knob=3,
            is_workday=WEEKDAYS,
            horizon_days=14,
        )
        is None
    )


def test_pack_pour_zero_tracks_returns_none() -> None:
    weeks = build_weeks(EXAMPLE_A_TODAY, {}, week_count=1, is_workday=WEEKDAYS)
    assert (
        pack_pour(0, {}, weeks, today=EXAMPLE_A_TODAY, knob=3, is_workday=WEEKDAYS)
        is None
    )


def test_build_quote_example_a_uses_pour_plan() -> None:
    weeks = build_weeks(
        EXAMPLE_A_TODAY, EXAMPLE_A_OCCUPANCY, week_count=4, is_workday=WEEKDAYS
    )
    quote = build_quote(
        5 * MAX_TRACK_LENGTH_M,
        weeks,
        today=EXAMPLE_A_TODAY,
        knob=3,
        occupancy=EXAMPLE_A_OCCUPANCY,
        is_workday=WEEKDAYS,
    )
    assert quote.tracks == 5
    assert quote.first_pour_date == date(2026, 9, 9)
    assert quote.first_pour_free == 2
    assert quote.solo_date == date(2026, 9, 10)
    assert quote.solo_week_end_date == date(2026, 9, 11)
    assert quote.earliest_start_week == WEEK_1
    assert quote.window is not None
    assert quote.window.promised_date == date(2026, 9, 11)
    assert quote.window.from_week == WEEK_1
    assert quote.window.to_week == WEEK_1
    assert quote.window.allocations == ((WEEK_1, 5),)


def test_module_has_no_app_sqlite_or_io() -> None:
    source = MODULE_PATH.read_text(encoding="utf-8")
    tree = ast.parse(source)
    forbidden_calls = {"load_holidays", "load_extra_workdays", "open"}
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert not alias.name.startswith("app"), alias.name
                assert alias.name != "sqlite3"
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            assert not module.startswith("app"), module
            assert module != "sqlite3"
        elif isinstance(node, ast.Call):
            func = node.func
            if isinstance(func, ast.Name):
                assert func.id not in forbidden_calls
            elif isinstance(func, ast.Attribute):
                assert func.attr not in forbidden_calls
