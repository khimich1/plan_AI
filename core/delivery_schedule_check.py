"""Светофор графика поставки: чистая проверка партий по загрузке производства.

Алгоритм (docs/specs/delivery-schedule.md):
1. Остаток партии = Σ max(0, qty − produced) по позициям.
2. Потребность в дорожках = ceil( Σ(остаток × length_m) / 101 × 1.15 ).
3. Симуляция от ``today`` по рабочим дням: свободно = max − occupied;
   партии в порядке ``produce_by`` съедают ёмкость.
4. Статус: green / yellow / red (+ подсказка дефицита для red).

Контракт ``workdays``
---------------------
Передаётся конечный ``set`` (или Iterable) ISO-дат рабочих дней.
День считается рабочим, если:
- его ISO есть в ``workdays``, или
- дата лежит строго вне диапазона ``[min(workdays), max(workdays)]``
  и это пн–пт (детерминированный fallback, чтобы короткий set в тестах
  не обрывал горизонт симуляции).

Дата внутри диапазона, но отсутствующая в set — нерабочий день
(выходной/праздник, который вызывающий код намеренно не передал).
Пустой set → все пн–пт от ``today``.

Риск R2 (occupancy)
-------------------
Дата вне ``occupancy`` = свободный день с ``occupied=0``,
``max=TRACKS_PER_DAY_DEFAULT``. Без I/O и без импортов ``app.*``.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Callable, Iterable, Literal

from core.production_capacity import MAX_TRACK_LENGTH_M, TRACKS_PER_DAY_DEFAULT

Status = Literal["green", "yellow", "red"]

# Страховка от бесконечного цикла при нулевой ёмкости на горизонте.
_MAX_SIMULATION_DAYS = 366 * 5


@dataclass(frozen=True)
class BatchItemInput:
    plate_id: int
    qty: int
    length_m: float


@dataclass(frozen=True)
class BatchInput:
    id: int | str
    name: str
    produce_by: str  # ISO YYYY-MM-DD
    items: list[BatchItemInput]


@dataclass(frozen=True)
class BatchCheck:
    batch_id: int | str
    status: Status
    ready_date: str | None  # ISO; None если потребность 0
    remaining_qty: int
    tracks_needed: int
    hint: str | None


def check_batches(
    *,
    batches: list[BatchInput],
    occupancy: dict[str, dict],
    workdays: set[str] | Iterable[str],
    produced: dict[int, int],
    today: str,
    capacity_buffer: float = 1.15,
    green_slack_workdays: int = 5,
) -> list[BatchCheck]:
    """Проверяет партии и возвращает светофор в порядке входного списка.

    Параметры
    ---------
    occupancy:
        date ISO → ``{occupied, max}``. Дата вне словаря — R2-дефолт.
    workdays:
        ISO рабочих дней; при нехватке горизонта — пн–пт (см. модуль).
    produced:
        ``plate_id → qty`` уже произведённое и привязанное к КП.
    today:
        ISO «сегодня» (для детерминизма тестов).
    """
    workday_set = set(workdays)
    is_workday = _make_workday_predicate(workday_set)
    today_d = date.fromisoformat(today)

    # Симуляция мутирует свободную ёмкость; ключ — ISO даты.
    free_by_day: dict[str, float] = {}

    def free_on(day: date) -> float:
        key = day.isoformat()
        if key not in free_by_day:
            free_by_day[key] = _day_free_capacity(key, occupancy)
        return free_by_day[key]

    def set_free(day: date, value: float) -> None:
        free_by_day[day.isoformat()] = value

    # Считаем метрики партий; симулируем в порядке produce_by, затем id.
    metrics: dict[int | str, tuple[int, int]] = {}
    for batch in batches:
        rem_qty, tracks = _batch_remaining_and_tracks(
            batch, produced, capacity_buffer=capacity_buffer
        )
        metrics[batch.id] = (rem_qty, tracks)

    ordered = sorted(
        batches,
        key=lambda b: (b.produce_by, str(b.id)),
    )
    results: dict[int | str, BatchCheck] = {}

    for batch in ordered:
        rem_qty, tracks_needed = metrics[batch.id]
        produce_by_d = date.fromisoformat(batch.produce_by)

        if tracks_needed <= 0:
            results[batch.id] = BatchCheck(
                batch_id=batch.id,
                status="green",
                ready_date=None,
                remaining_qty=rem_qty,
                tracks_needed=0,
                hint=None,
            )
            continue

        # Дефицит для red-hint: ёмкость в окне today→produce_by до съедания
        # этой партией (ранее обработанные партии уже уменьшили free_by_day).
        available_before = _sum_free_in_window(
            start=today_d,
            end=produce_by_d,
            is_workday=is_workday,
            free_on=free_on,
        )
        ready_d = _simulate_ready_date(
            tracks_needed=tracks_needed,
            start=today_d,
            is_workday=is_workday,
            free_on=free_on,
            set_free=set_free,
        )

        status, hint = _status_and_hint(
            ready_d=ready_d,
            produce_by_d=produce_by_d,
            tracks_needed=tracks_needed,
            available_before=available_before,
            is_workday=is_workday,
            green_slack_workdays=green_slack_workdays,
        )
        results[batch.id] = BatchCheck(
            batch_id=batch.id,
            status=status,
            ready_date=ready_d.isoformat(),
            remaining_qty=rem_qty,
            tracks_needed=tracks_needed,
            hint=hint,
        )

    return [results[b.id] for b in batches]


def _batch_remaining_and_tracks(
    batch: BatchInput,
    produced: dict[int, int],
    *,
    capacity_buffer: float,
) -> tuple[int, int]:
    remaining_qty = 0
    remaining_length_m = 0.0
    for item in batch.items:
        rem = max(0, int(item.qty) - int(produced.get(item.plate_id, 0)))
        remaining_qty += rem
        remaining_length_m += rem * float(item.length_m)

    if remaining_length_m <= 0:
        return remaining_qty, 0

    raw_tracks = (remaining_length_m / MAX_TRACK_LENGTH_M) * capacity_buffer
    tracks_needed = int(math.ceil(raw_tracks - 1e-12))
    return remaining_qty, max(0, tracks_needed)


def _day_free_capacity(day_iso: str, occupancy: dict[str, dict]) -> float:
    """Свободные дорожки на дату; вне occupancy — R2-дефолт."""
    info = occupancy.get(day_iso)
    if info is None:
        return float(TRACKS_PER_DAY_DEFAULT)
    max_tracks = float(info.get("max", TRACKS_PER_DAY_DEFAULT))
    occupied = float(info.get("occupied", 0))
    return max(0.0, max_tracks - occupied)


def _make_workday_predicate(workdays: set[str]) -> Callable[[date], bool]:
    if workdays:
        lo = min(date.fromisoformat(d) for d in workdays)
        hi = max(date.fromisoformat(d) for d in workdays)
    else:
        lo = hi = None

    def is_workday(d: date) -> bool:
        key = d.isoformat()
        if key in workdays:
            return True
        # Вне известного календаря — детерминированный пн–пт.
        if lo is None or d < lo or d > hi:
            return d.weekday() < 5
        return False

    return is_workday


def _iter_workdays(
    start: date,
    is_workday: Callable[[date], bool],
) -> Iterable[date]:
    current = start
    for _ in range(_MAX_SIMULATION_DAYS):
        if is_workday(current):
            yield current
        current += timedelta(days=1)


def _sum_free_in_window(
    *,
    start: date,
    end: date,
    is_workday: Callable[[date], bool],
    free_on: Callable[[date], float],
) -> float:
    """Сумма свободной ёмкости на рабочих днях start..end включительно."""
    if end < start:
        return 0.0
    total = 0.0
    current = start
    # Ограничиваем обход окном + небольшой запас (end может быть далеко).
    guard = (end - start).days + 2
    for _ in range(max(guard, 1)):
        if current > end:
            break
        if is_workday(current):
            total += free_on(current)
        current += timedelta(days=1)
    return total


def _simulate_ready_date(
    *,
    tracks_needed: int,
    start: date,
    is_workday: Callable[[date], bool],
    free_on: Callable[[date], float],
    set_free: Callable[[date, float], None],
) -> date:
    """Съедает ``tracks_needed`` дорожек с ``start``; возвращает день готовности."""
    remaining = float(tracks_needed)
    last_day = start
    for day in _iter_workdays(start, is_workday):
        last_day = day
        free = free_on(day)
        if free <= 0:
            continue
        take = min(free, remaining)
        set_free(day, free - take)
        remaining -= take
        if remaining <= 1e-12:
            return day
    # Горизонт исчерпан — фиксируем последний рассмотренный день.
    return last_day


def _shift_workdays_back(
    from_day: date,
    n: int,
    is_workday: Callable[[date], bool],
) -> date:
    """Дата на ``n`` рабочих дней раньше ``from_day`` (сам from_day не считается)."""
    if n <= 0:
        return from_day
    current = from_day
    left = n
    # Достаточный запас назад (включая длинные праздники).
    for _ in range(_MAX_SIMULATION_DAYS):
        current -= timedelta(days=1)
        if is_workday(current):
            left -= 1
            if left == 0:
                return current
    return current


def _status_and_hint(
    *,
    ready_d: date,
    produce_by_d: date,
    tracks_needed: int,
    available_before: float,
    is_workday: Callable[[date], bool],
    green_slack_workdays: int,
) -> tuple[Status, str | None]:
    green_deadline = _shift_workdays_back(
        produce_by_d, green_slack_workdays, is_workday
    )
    if ready_d <= green_deadline:
        return "green", None
    if ready_d <= produce_by_d:
        return "yellow", None

    # Красный: дефицит ёмкости в окне today→produce_by (до потребления партии).
    deficit = float(tracks_needed) - available_before
    n_extra = int(math.ceil(deficit - 1e-12)) if deficit > 1e-12 else 0
    if n_extra < 1:
        # Готовность после дедлайна при «нулевом» дефиците (граничные даты) —
        # всё равно подсказываем минимум +1 дорожку.
        n_extra = 1
    hint = (
        f"нужно +{n_extra} дорожек до {produce_by_d.strftime('%d.%m.%Y')}"
    )
    return "red", hint
