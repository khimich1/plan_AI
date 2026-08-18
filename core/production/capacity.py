"""Pure day-capacity helpers for production planning (no I/O, no app imports)."""

from __future__ import annotations

import math
from collections.abc import Callable, Collection, Mapping, Sequence
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, Literal

from core.production.errors import PlanBuildError
from core.production_capacity import MAX_TRACK_LENGTH_M, TRACKS_PER_DAY_DEFAULT

# Hard factory cap: never more than 5 tracks/day. Floor 0 = day manually disabled.
TRACKS_PER_DAY_HARD_CAP = 5
MAX_DEFICIT_OPTIONS = 10
FUTURE_HORIZON_DAYS = 30

CapacityOptionAction = Literal["bump_fill", "propose_day"]


@dataclass(frozen=True, slots=True)
class CapacityOption:
    action: CapacityOptionAction
    date: str  # ISO YYYY-MM-DD
    add_tracks: int
    free: int


@dataclass(frozen=True, slots=True)
class CapacityDeficit:
    tracks_needed: int
    tracks_available: int
    tracks_missing: int
    deficit_until: str  # ISO
    options: tuple[CapacityOption, ...]


def clamp_day_max(value: int) -> int:
    """Clamp a day max to the allowed range ``0…TRACKS_PER_DAY_HARD_CAP``."""
    return max(0, min(int(value), TRACKS_PER_DAY_HARD_CAP))


def _to_iso_date(value: date | str) -> str:
    if isinstance(value, date):
        return value.isoformat()
    return date.fromisoformat(str(value)).isoformat()


def _normalize_overrides(
    overrides: Mapping[date | str, int],
) -> dict[str, int]:
    return {
        _to_iso_date(key): clamp_day_max(int(value)) for key, value in overrides.items()
    }


def get_day_capacity(
    day: date | str,
    overrides: Mapping[date | str, int],
    *,
    default: int | None = None,
) -> int:
    """Return max tracks for ``day`` from overrides, else ``default`` (always ≤ hard cap)."""
    day_iso = _to_iso_date(day)
    normalized = _normalize_overrides(overrides)
    if day_iso in normalized:
        return int(normalized[day_iso])
    if default is None:
        return clamp_day_max(int(TRACKS_PER_DAY_DEFAULT))
    return clamp_day_max(int(default))


def _day_max_tracks(day_iso: str, day_capacity: Mapping[str, int]) -> int:
    if day_iso in day_capacity:
        return clamp_day_max(int(day_capacity[day_iso]))
    return get_day_capacity(day_iso, {})


def day_free_tracks(
    day: date | str,
    day_capacity: Mapping[str, int],
    occupancy: Mapping[str, int] | None = None,
) -> int:
    """``free = day_max − occupied`` with ``day_max ≤ TRACKS_PER_DAY_HARD_CAP``."""
    day_iso = _to_iso_date(day)
    day_max = _day_max_tracks(day_iso, day_capacity)
    occupied = int((occupancy or {}).get(day_iso, 0) or 0)
    return max(0, day_max - occupied)


def validate_fill_targets(
    fill_targets: Sequence[Mapping[str, Any]],
    day_capacity: Mapping[str, int],
    occupancy: Mapping[str, int] | None = None,
) -> None:
    """Validate fill_targets against free slots (``day_max − occupied``).

    ``occupancy`` defaults to empty (all slots free). Raises ``PlanBuildError``
    when targets are empty, malformed, or request more than free tracks.
    """
    if not fill_targets:
        raise PlanBuildError("fill_targets пуст.")

    occ = occupancy or {}
    for target in fill_targets:
        day_raw = target.get("date") or ""
        try:
            day_iso = _to_iso_date(str(day_raw))
        except ValueError as exc:
            raise PlanBuildError(
                f"Неверный формат даты в fill_targets: {day_raw}"
            ) from exc

        tracks_raw = target.get("tracks", 0)
        try:
            tracks = int(tracks_raw)
        except (TypeError, ValueError) as exc:
            raise PlanBuildError(
                f"На {day_iso} запрошено недопустимое число дорожек: {tracks_raw!r}."
            ) from exc

        if tracks < 1:
            raise PlanBuildError(
                f"На {day_iso} запрошено {tracks} дорожек — должно быть >= 1."
            )

        free = day_free_tracks(day_iso, day_capacity, occ)
        if tracks > free:
            raise PlanBuildError(
                f"На {day_iso} свободно {free} дорожек, запрошено {tracks}."
            )


def _default_is_workday(day: date) -> bool:
    return day.weekday() < 5


def calculate_capacity_deficit(
    urgent_length_m: float,
    fill_targets: Sequence[Mapping[str, Any]],
    day_capacity: Mapping[str, int],
    *,
    occupancy: Mapping[str, int] | None = None,
    completed_dates: Collection[str] | None = None,
    today: date | str | None = None,
    is_workday: Callable[[date], bool] | None = None,
    max_options: int = MAX_DEFICIT_OPTIONS,
    future_horizon_days: int = FUTURE_HORIZON_DAYS,
) -> CapacityDeficit | None:
    """Compute capacity deficit and ordered refill options (A→B→C).

    Options are suggestions only — caller/UI applies them; never mutate capacity.
    """
    if urgent_length_m <= 0:
        tracks_needed = 0
    else:
        tracks_needed = int(math.ceil(urgent_length_m / MAX_TRACK_LENGTH_M))

    tracks_available = sum(int(ft["tracks"]) for ft in fill_targets)
    tracks_missing = max(0, tracks_needed - tracks_available)
    if tracks_missing == 0:
        return None

    if not fill_targets:
        return None

    dated_targets: list[tuple[str, int]] = []
    for ft in fill_targets:
        day_iso = _to_iso_date(str(ft["date"]))
        dated_targets.append((day_iso, int(ft["tracks"])))
    dated_targets.sort(key=lambda item: item[0])

    deficit_until = dated_targets[-1][0]
    occ = occupancy or {}
    completed = {_to_iso_date(d) for d in (completed_dates or ())}
    workday_fn = is_workday or _default_is_workday
    today_d = date.fromisoformat(_to_iso_date(today)) if today is not None else date.today()

    selected_dates = {day_iso for day_iso, _ in dated_targets}
    min_selected = date.fromisoformat(dated_targets[0][0])
    max_selected = date.fromisoformat(dated_targets[-1][0])

    options: list[CapacityOption] = []

    # Step A — bump fill in selected days with headroom
    for day_iso, tracks in dated_targets:
        free = day_free_tracks(day_iso, day_capacity, occ)
        headroom = free - tracks
        if headroom > 0:
            options.append(
                CapacityOption(
                    action="bump_fill",
                    date=day_iso,
                    add_tracks=headroom,
                    free=free,
                )
            )

    # Step B — previous days: today ≤ date < min(fill_targets)
    if len(options) < max_options:
        cursor = today_d
        while cursor < min_selected and len(options) < max_options:
            day_iso = cursor.isoformat()
            if (
                day_iso not in selected_dates
                and day_iso not in completed
                and workday_fn(cursor)
            ):
                day_max = _day_max_tracks(day_iso, day_capacity)
                free = day_free_tracks(day_iso, day_capacity, occ)
                if day_max > 0 and free > 0:
                    options.append(
                        CapacityOption(
                            action="propose_day",
                            date=day_iso,
                            add_tracks=free,
                            free=free,
                        )
                    )
            cursor += timedelta(days=1)

    # Step C — future days after max(fill_targets), horizon calendar days
    if len(options) < max_options:
        horizon_end = max_selected + timedelta(days=future_horizon_days)
        cursor = max_selected + timedelta(days=1)
        while cursor <= horizon_end and len(options) < max_options:
            day_iso = cursor.isoformat()
            if (
                day_iso not in selected_dates
                and day_iso not in completed
                and workday_fn(cursor)
            ):
                day_max = _day_max_tracks(day_iso, day_capacity)
                free = day_free_tracks(day_iso, day_capacity, occ)
                if day_max > 0 and free > 0:
                    options.append(
                        CapacityOption(
                            action="propose_day",
                            date=day_iso,
                            add_tracks=free,
                            free=free,
                        )
                    )
            cursor += timedelta(days=1)

    return CapacityDeficit(
        tracks_needed=tracks_needed,
        tracks_available=tracks_available,
        tracks_missing=tracks_missing,
        deficit_until=deficit_until,
        options=tuple(options[:max_options]),
    )
