"""Unit-тесты capacity gate: snapshot math + enforce (red block / yellow allow)."""

from __future__ import annotations

from datetime import date, timedelta

import pytest

from core.delivery_schedule_check import BatchInput, BatchItemInput
from core.production_capacity import MAX_TRACK_LENGTH_M, TRACKS_PER_DAY_DEFAULT
from app.services.capacity_gate_service import (
    CapacityGateBlockedError,
    CapacitySnapshot,
    assert_capacity_allows_save,
    build_capacity_snapshot,
    build_multi_batch_gate,
)

_UNIT_LENGTH_M = MAX_TRACK_LENGTH_M / 1.15
_TODAY = "2026-03-02"  # Mon
_TOMORROW = "2026-03-03"


def _item(plate_id: int, qty: int) -> BatchItemInput:
    return BatchItemInput(plate_id=plate_id, qty=qty, length_m=_UNIT_LENGTH_M)


def _full_capacity(start: str, n_days: int = 40) -> dict[str, dict]:
    start_d = date.fromisoformat(start)
    out: dict[str, dict] = {}
    for i in range(n_days):
        key = (start_d + timedelta(days=i)).isoformat()
        out[key] = {"occupied": 0, "max": TRACKS_PER_DAY_DEFAULT}
    return out


def _mon_fri(start: str, n_days: int = 40) -> set[str]:
    start_d = date.fromisoformat(start)
    out: set[str] = set()
    for i in range(n_days):
        d = start_d + timedelta(days=i)
        if d.weekday() < 5:
            out.add(d.isoformat())
    return out


def test_snapshot_green_with_slack() -> None:
    snap = build_capacity_snapshot(
        items=[_item(1, 1)],
        target_date="2026-03-20",
        occupancy=_full_capacity(_TODAY),
        workdays=_mon_fri(_TODAY),
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    assert isinstance(snap, CapacitySnapshot)
    assert snap.start_date == _TOMORROW
    assert snap.target_date == "2026-03-20"
    assert snap.tracks_needed == 1
    assert snap.tracks_free_in_window >= 1
    assert snap.delta == snap.tracks_free_in_window - snap.tracks_needed
    assert snap.status == "green"
    assert snap.hint is None
    assert snap.calendar_from_month == "2026-03"
    assert snap.calendar_to_month == "2026-03"


def test_snapshot_red_on_deficit() -> None:
    # target = завтра: окно 1 раб.день × 5 дорожек, нужно 20 → red.
    snap = build_capacity_snapshot(
        items=[_item(1, 20)],
        target_date=_TOMORROW,
        occupancy=_full_capacity(_TODAY),
        workdays=_mon_fri(_TODAY),
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    assert snap.status == "red"
    assert snap.tracks_needed == 20
    assert snap.hint is not None
    assert "нужно" in snap.hint
    assert snap.delta < 0


def test_snapshot_yellow_near_deadline() -> None:
    # 1 дорожка, target = завтра → ready = завтра, slack 5 → yellow.
    snap = build_capacity_snapshot(
        items=[_item(1, 1)],
        target_date=_TOMORROW,
        occupancy=_full_capacity(_TODAY),
        workdays=_mon_fri(_TODAY),
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    assert snap.status == "yellow"
    assert snap.hint is None


def test_snapshot_calendar_months_span() -> None:
    snap = build_capacity_snapshot(
        items=[_item(1, 1)],
        target_date="2026-05-15",
        occupancy=_full_capacity(_TODAY, n_days=90),
        workdays=_mon_fri(_TODAY, n_days=90),
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    assert snap.calendar_from_month == "2026-03"
    assert snap.calendar_to_month == "2026-05"


def test_snapshot_default_start_is_tomorrow() -> None:
    """Без start_date — старт с завтра относительно today."""
    snap = build_capacity_snapshot(
        items=[_item(1, 1)],
        target_date="2026-03-20",
        occupancy=_full_capacity(_TODAY),
        workdays=_mon_fri(_TODAY),
        produced={},
        today=_TODAY,
    )
    assert snap.start_date == _TOMORROW


def test_assert_blocks_red_allows_yellow_green() -> None:
    red = build_capacity_snapshot(
        items=[_item(1, 20)],
        target_date=_TOMORROW,
        occupancy=_full_capacity(_TODAY),
        workdays=_mon_fri(_TODAY),
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    with pytest.raises(CapacityGateBlockedError) as exc_info:
        assert_capacity_allows_save(red)
    assert "нужно" in str(exc_info.value)

    yellow = build_capacity_snapshot(
        items=[_item(1, 1)],
        target_date=_TOMORROW,
        occupancy=_full_capacity(_TODAY),
        workdays=_mon_fri(_TODAY),
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    assert_capacity_allows_save(yellow)  # no raise

    green = build_capacity_snapshot(
        items=[_item(1, 1)],
        target_date="2026-03-20",
        occupancy=_full_capacity(_TODAY),
        workdays=_mon_fri(_TODAY),
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    assert_capacity_allows_save(green)


def test_multi_batch_gate_blocks_on_any_red() -> None:
    occupancy = {_TOMORROW: {"occupied": 0, "max": 5}}
    workdays = {_TOMORROW, "2026-03-04", "2026-03-05", "2026-03-06"}
    batches = [
        BatchInput(
            id=1,
            name="A",
            produce_by=_TOMORROW,
            items=[_item(1, 5)],
        ),
        BatchInput(
            id=2,
            name="B",
            produce_by=_TOMORROW,
            items=[_item(2, 5)],
        ),
    ]
    gate = build_multi_batch_gate(
        batches=batches,
        occupancy=occupancy,
        workdays=workdays,
        produced={},
        today=_TODAY,
        start_date=_TOMORROW,
    )
    assert gate.worst_status == "red"
    assert gate.blocking_hint is not None
    with pytest.raises(CapacityGateBlockedError):
        assert_capacity_allows_save(gate)
