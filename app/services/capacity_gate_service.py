"""Ёмкость завода: snapshot для виджета + жёсткий гейт при status=red.

Математика статуса — через ``core.delivery_schedule_check.check_batches``
(те же пороги: buffer 1.15, slack 5 раб.дней, 101 м, 5 дор/день).
Старт окна по умолчанию — завтра относительно ``today``.
Занятость — только план (``occupancy`` / days_info); СГП не учитывается.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Literal

from core.delivery_schedule_check import (
    BatchCheck,
    BatchInput,
    BatchItemInput,
    check_batches,
    _day_free_capacity,
    _make_workday_predicate,
    _sum_free_in_window,
)

Status = Literal["green", "yellow", "red"]

_STATUS_RANK: dict[Status, int] = {"green": 0, "yellow": 1, "red": 2}


class CapacityGateBlockedError(Exception):
    """Red-статус: сохранение / перевод в производство запрещены."""

    def __init__(self, message: str, *, hint: str | None = None) -> None:
        super().__init__(message)
        self.hint = hint


@dataclass(frozen=True)
class CapacitySnapshot:
    start_date: str
    target_date: str
    tracks_needed: int
    tracks_free_in_window: int
    delta: int
    status: Status
    hint: str | None
    days_info: dict[str, dict]
    holidays: list[str]
    extra_workdays: list[str]
    calendar_from_month: str
    calendar_to_month: str


@dataclass(frozen=True)
class MultiBatchGateResult:
    """Агрегат по нескольким партиям: блок если любой red."""

    worst_status: Status
    blocking_hint: str | None
    checks: list[BatchCheck]
    start_date: str
    target_date: str
    tracks_needed: int
    tracks_free_in_window: int
    delta: int
    days_info: dict[str, dict]
    holidays: list[str]
    extra_workdays: list[str]
    calendar_from_month: str
    calendar_to_month: str

    @property
    def status(self) -> Status:
        return self.worst_status

    @property
    def hint(self) -> str | None:
        return self.blocking_hint


def tomorrow_iso(today: str) -> str:
    return (date.fromisoformat(today) + timedelta(days=1)).isoformat()


def resolve_start_date(*, today: str, start_date: str | None = None) -> str:
    if start_date is not None:
        return date.fromisoformat(start_date).isoformat()
    return tomorrow_iso(today)


def build_capacity_snapshot(
    *,
    items: list[BatchItemInput],
    target_date: str,
    occupancy: dict[str, dict],
    workdays: set[str] | list[str],
    produced: dict[int, int],
    today: str,
    start_date: str | None = None,
    holidays: list[str] | None = None,
    extra_workdays: list[str] | None = None,
    capacity_buffer: float = 1.15,
    green_slack_workdays: int = 5,
) -> CapacitySnapshot:
    """КП как одна виртуальная партия с ``produce_by=target_date``."""
    start = resolve_start_date(today=today, start_date=start_date)
    target = date.fromisoformat(target_date).isoformat()
    workday_set = set(workdays)

    batch = BatchInput(
        id="kp",
        name="КП",
        produce_by=target,
        items=list(items),
    )
    checks = check_batches(
        batches=[batch],
        occupancy=occupancy,
        workdays=workday_set,
        produced=produced,
        today=today,
        start_date=start,
        capacity_buffer=capacity_buffer,
        green_slack_workdays=green_slack_workdays,
    )
    check = checks[0]

    free = _tracks_free_in_window(
        start_iso=start,
        end_iso=target,
        occupancy=occupancy,
        workdays=workday_set,
    )
    needed = int(check.tracks_needed)
    today_d = date.fromisoformat(today)
    target_d = date.fromisoformat(target)

    return CapacitySnapshot(
        start_date=start,
        target_date=target,
        tracks_needed=needed,
        tracks_free_in_window=int(free),
        delta=int(free) - needed,
        status=check.status,
        hint=check.hint,
        days_info=dict(occupancy),
        holidays=list(holidays or []),
        extra_workdays=list(extra_workdays or []),
        calendar_from_month=f"{today_d.year:04d}-{today_d.month:02d}",
        calendar_to_month=f"{target_d.year:04d}-{target_d.month:02d}",
    )


def build_multi_batch_gate(
    *,
    batches: list[BatchInput],
    occupancy: dict[str, dict],
    workdays: set[str] | list[str],
    produced: dict[int, int],
    today: str,
    start_date: str | None = None,
    holidays: list[str] | None = None,
    extra_workdays: list[str] | None = None,
    capacity_buffer: float = 1.15,
    green_slack_workdays: int = 5,
) -> MultiBatchGateResult:
    """Светофор по партиям; worst = max(status); блок при любом red."""
    start = resolve_start_date(today=today, start_date=start_date)
    workday_set = set(workdays)
    checks = check_batches(
        batches=batches,
        occupancy=occupancy,
        workdays=workday_set,
        produced=produced,
        today=today,
        start_date=start,
        capacity_buffer=capacity_buffer,
        green_slack_workdays=green_slack_workdays,
    )

    worst: Status = "green"
    blocking_hint: str | None = None
    for check in checks:
        if _STATUS_RANK[check.status] > _STATUS_RANK[worst]:
            worst = check.status
        if check.status == "red" and blocking_hint is None:
            blocking_hint = check.hint

    target = max((b.produce_by for b in batches), default=start)
    free = _tracks_free_in_window(
        start_iso=start,
        end_iso=target,
        occupancy=occupancy,
        workdays=workday_set,
    )
    needed = sum(int(c.tracks_needed) for c in checks)
    today_d = date.fromisoformat(today)
    target_d = date.fromisoformat(target)

    return MultiBatchGateResult(
        worst_status=worst,
        blocking_hint=blocking_hint,
        checks=checks,
        start_date=start,
        target_date=target,
        tracks_needed=needed,
        tracks_free_in_window=int(free),
        delta=int(free) - needed,
        days_info=dict(occupancy),
        holidays=list(holidays or []),
        extra_workdays=list(extra_workdays or []),
        calendar_from_month=f"{today_d.year:04d}-{today_d.month:02d}",
        calendar_to_month=f"{target_d.year:04d}-{target_d.month:02d}",
    )


def assert_capacity_allows_save(
    snapshot: CapacitySnapshot | MultiBatchGateResult,
) -> None:
    """Yellow/green — ok; red → CapacityGateBlockedError с hint."""
    if snapshot.status != "red":
        return
    hint = snapshot.hint
    message = hint or (
        "Завод перегружен до выбранного срока. "
        "Увеличьте срок изготовления и повторите."
    )
    raise CapacityGateBlockedError(message, hint=hint)


def _tracks_free_in_window(
    *,
    start_iso: str,
    end_iso: str,
    occupancy: dict[str, dict],
    workdays: set[str],
) -> float:
    start_d = date.fromisoformat(start_iso)
    end_d = date.fromisoformat(end_iso)
    is_workday = _make_workday_predicate(workdays)

    def free_on(day: date) -> float:
        return _day_free_capacity(day.isoformat(), occupancy)

    return _sum_free_in_window(
        start=start_d,
        end=end_d,
        is_workday=is_workday,
        free_on=free_on,
    )
