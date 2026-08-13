"""Unit-тесты светофора графика поставки (core/delivery_schedule_check).

Детерминированные кейсы Acceptance из docs/specs/delivery-schedule.md:
фиксированные today / occupancy / workdays, без БД и сети.
"""

from __future__ import annotations

import re
from datetime import date, timedelta

import pytest

from core.delivery_schedule_check import (
    BatchCheck,
    BatchInput,
    BatchItemInput,
    check_batches,
)
from core.production_capacity import MAX_TRACK_LENGTH_M, TRACKS_PER_DAY_DEFAULT

# Длина плиты: 1 шт. × buffer → ровно 1 дорожка (ceil(1.0 − eps)).
_UNIT_LENGTH_M = MAX_TRACK_LENGTH_M / 1.15

# Понедельник — удобная точка отсчёта для пн–пт.
_TODAY = "2026-03-02"


def _item(plate_id: int, qty: int, length_m: float = _UNIT_LENGTH_M) -> BatchItemInput:
    return BatchItemInput(plate_id=plate_id, qty=qty, length_m=length_m)


def _batch(
    batch_id: int | str,
    produce_by: str,
    qty: int,
    *,
    plate_id: int = 1,
    name: str | None = None,
) -> BatchInput:
    return BatchInput(
        id=batch_id,
        name=name or f"batch-{batch_id}",
        produce_by=produce_by,
        items=[_item(plate_id, qty)],
    )


def _full_capacity_days(start: str, n_days: int = 40) -> dict[str, dict]:
    """occupancy с max=TRACKS_PER_DAY_DEFAULT и occupied=0 на n календарных дней."""
    start_d = date.fromisoformat(start)
    out: dict[str, dict] = {}
    for i in range(n_days):
        key = (start_d + timedelta(days=i)).isoformat()
        out[key] = {"occupied": 0, "max": TRACKS_PER_DAY_DEFAULT}
    return out


def _mon_fri_workdays(start: str, n_days: int = 40) -> set[str]:
    start_d = date.fromisoformat(start)
    out: set[str] = set()
    for i in range(n_days):
        d = start_d + timedelta(days=i)
        if d.weekday() < 5:
            out.add(d.isoformat())
    return out


def test_status_green_when_ready_has_slack() -> None:
    """ready ≪ produce_by − 5 раб.дней → green."""
    # 1 дорожка → ready = today; produce_by далеко вперёд.
    batches = [_batch(1, "2026-03-20", qty=1)]
    results = check_batches(
        batches=batches,
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={},
        today=_TODAY,
    )
    assert len(results) == 1
    assert results[0].status == "green"
    assert results[0].ready_date == _TODAY
    assert results[0].tracks_needed == 1
    assert results[0].remaining_qty == 1
    assert results[0].hint is None


def test_status_yellow_when_ready_within_deadline_but_no_slack() -> None:
    """ready ≤ produce_by, но позже green-deadline → yellow."""
    # produce_by = today: 1 дорожка готова сегодня → yellow (slack = 5 раб.дней назад).
    batches = [_batch(1, _TODAY, qty=1)]
    results = check_batches(
        batches=batches,
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={},
        today=_TODAY,
    )
    assert results[0].status == "yellow"
    assert results[0].ready_date == _TODAY
    assert results[0].hint is None


def test_status_red_when_ready_after_produce_by() -> None:
    """ready > produce_by → red + hint дефицита."""
    # 15 дорожек при 5/день → ready минимум через 3 раб.дня; produce_by = today.
    batches = [_batch(1, _TODAY, qty=15)]
    results = check_batches(
        batches=batches,
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={},
        today=_TODAY,
    )
    assert results[0].status == "red"
    assert results[0].ready_date is not None
    assert results[0].ready_date > _TODAY
    assert results[0].tracks_needed == 15
    assert results[0].hint is not None
    assert "нужно" in results[0].hint


def test_remaining_qty_subtracts_produced() -> None:
    batches = [_batch(1, "2026-03-20", qty=10, plate_id=7)]
    results = check_batches(
        batches=batches,
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={7: 3},
        today=_TODAY,
    )
    assert results[0].remaining_qty == 7
    assert results[0].tracks_needed == 7


def test_fully_produced_batch_is_green_with_zero_tracks() -> None:
    batches = [_batch(1, _TODAY, qty=10, plate_id=7)]
    results = check_batches(
        batches=batches,
        occupancy={},
        workdays=set(),
        produced={7: 10},
        today=_TODAY,
    )
    assert results[0].status == "green"
    assert results[0].remaining_qty == 0
    assert results[0].tracks_needed == 0
    assert results[0].ready_date is None
    assert results[0].hint is None


def test_produced_above_qty_clamps_remaining_to_zero() -> None:
    batches = [_batch(1, "2026-03-20", qty=5, plate_id=3)]
    results = check_batches(
        batches=batches,
        occupancy={},
        workdays=set(),
        produced={3: 99},
        today=_TODAY,
    )
    assert results[0].remaining_qty == 0
    assert results[0].tracks_needed == 0
    assert results[0].status == "green"
    assert results[0].ready_date is None


def test_weekends_skipped_via_workdays_set_gap() -> None:
    """В диапазоне set дата отсутствует → выходной; симуляция перескакивает."""
    # today = пт 2026-03-06; сб/вс нет в set; пн 2026-03-09 есть.
    today = "2026-03-06"
    # 6 дорожек: 5 в пт + 1 в пн → ready = понедельник.
    batches = [_batch(1, "2026-03-20", qty=6)]
    workdays = {
        "2026-03-06",  # Fri
        "2026-03-09",  # Mon
        "2026-03-10",
        "2026-03-11",
        "2026-03-12",
        "2026-03-13",
        "2026-03-16",
        "2026-03-17",
        "2026-03-18",
        "2026-03-19",
        "2026-03-20",
    }
    results = check_batches(
        batches=batches,
        occupancy=_full_capacity_days(today),
        workdays=workdays,
        produced={},
        today=today,
    )
    assert results[0].ready_date == "2026-03-09"
    assert results[0].status == "green"


def test_weekends_skipped_outside_workdays_range_mon_fri_fallback() -> None:
    """Пустой workdays → пн–пт; сб/вс не рабочие."""
    today = "2026-03-06"  # Friday
    batches = [_batch(1, "2026-03-20", qty=6)]
    results = check_batches(
        batches=batches,
        occupancy=_full_capacity_days(today),
        workdays=set(),
        produced={},
        today=today,
    )
    assert results[0].ready_date == "2026-03-09"  # Monday after weekend


def test_two_batches_compete_later_turns_red() -> None:
    """Ранний produce_by съедает слоты; вторая партия в том же окне краснеет."""
    today = _TODAY
    # Окно today→produce_by: только сегодня, 5 свободных дорожек.
    produce_by = today
    occupancy = {today: {"occupied": 0, "max": 5}}
    # A раньше по produce_by (одинаковый дедлайн — tie-break по id), qty=5.
    # B позже в порядке симуляции (больший id), qty=5 — ёмкости уже нет.
    batch_a = _batch(1, produce_by, qty=5, plate_id=1)
    batch_b = _batch(2, produce_by, qty=5, plate_id=2)
    # Входной порядок: B, A — симуляция всё равно A затем B.
    results = check_batches(
        batches=[batch_b, batch_a],
        occupancy=occupancy,
        workdays={today},
        produced={},
        today=today,
    )
    by_id = {r.batch_id: r for r in results}
    assert by_id[1].status in ("green", "yellow")
    assert by_id[1].tracks_needed == 5
    assert by_id[2].status == "red"
    assert by_id[2].hint is not None


def test_r2_date_outside_occupancy_uses_default_tracks() -> None:
    """Дата вне occupancy → свободный день с max=TRACKS_PER_DAY_DEFAULT."""
    today = _TODAY
    # Occupancy пустой: весь горизонт — R2-дефолт (5 дорожек/день).
    batches = [_batch(1, "2026-03-20", qty=5)]
    results = check_batches(
        batches=batches,
        occupancy={},
        workdays=_mon_fri_workdays(today),
        produced={},
        today=today,
    )
    assert results[0].tracks_needed == 5
    # 5 дорожек за один дефолтный день → ready = today.
    assert results[0].ready_date == today
    assert results[0].status == "green"

    # Контроль: если «сегодня» полностью занят в occupancy, а завтра вне словаря —
    # завтра даёт 5 по R2.
    occupancy_today_full = {today: {"occupied": 5, "max": 5}}
    results2 = check_batches(
        batches=[_batch(1, "2026-03-20", qty=5)],
        occupancy=occupancy_today_full,
        workdays=_mon_fri_workdays(today),
        produced={},
        today=today,
    )
    assert results2[0].ready_date == "2026-03-03"  # Tue, вне occupancy → 5


def test_red_hint_contains_nuzhno_and_dd_mm_yyyy() -> None:
    # produce_by = today: окно 1 день × 5 дорожек, потребность 20 → red.
    produce_by = _TODAY
    batches = [_batch(1, produce_by, qty=20)]
    results = check_batches(
        batches=batches,
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={},
        today=_TODAY,
    )
    assert results[0].status == "red"
    hint = results[0].hint
    assert hint is not None
    assert "нужно" in hint
    assert re.search(r"\d{2}\.\d{2}\.\d{4}", hint), hint
    assert "02.03.2026" in hint


def test_result_order_matches_input_order_not_simulation_order() -> None:
    """Результат в порядке входного batches, даже если симуляция сортирует иначе."""
    # produce_by у B раньше → симуляция: B затем A; вход: A, B.
    batch_a = _batch("late", "2026-03-20", qty=1, plate_id=1)
    batch_b = _batch("early", "2026-03-10", qty=1, plate_id=2)
    results = check_batches(
        batches=[batch_a, batch_b],
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={},
        today=_TODAY,
    )
    assert [r.batch_id for r in results] == ["late", "early"]


@pytest.mark.parametrize(
    ("qty", "produce_by", "expected_status"),
    [
        (1, "2026-03-20", "green"),
        (1, _TODAY, "yellow"),
        (15, _TODAY, "red"),
    ],
)
def test_status_parametrize(
    qty: int,
    produce_by: str,
    expected_status: str,
) -> None:
    results = check_batches(
        batches=[_batch(1, produce_by, qty=qty)],
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={},
        today=_TODAY,
    )
    assert results[0].status == expected_status


def test_batch_check_fields_populated() -> None:
    results = check_batches(
        batches=[_batch(42, "2026-03-20", qty=2)],
        occupancy=_full_capacity_days(_TODAY),
        workdays=_mon_fri_workdays(_TODAY),
        produced={},
        today=_TODAY,
    )
    r = results[0]
    assert isinstance(r, BatchCheck)
    assert r.batch_id == 42
    assert r.remaining_qty == 2
    assert r.tracks_needed == 2
    assert r.ready_date == _TODAY
