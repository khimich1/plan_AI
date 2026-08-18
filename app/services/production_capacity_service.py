"""Service layer for per-day production capacity (overrides + default)."""

from __future__ import annotations

from datetime import date
from typing import Any, Mapping, Sequence

from app.core.settings import get_settings
from app.repositories.day_capacity_repository import DayCapacityRepository
from core.production.capacity import get_day_capacity
from core.production.capacity import validate_fill_targets as validate_capacity_fill_targets
from core.production.errors import PlanBuildError


class ProductionCapacityError(RuntimeError):
    """Domain error for capacity lookup, override, or fill_targets validation."""


def _to_date(value: date | str) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


def _extract_updated_by(user: Mapping[str, Any] | None) -> str | None:
    if not user:
        return None
    for key in ("email", "login", "user_id"):
        raw = user.get(key)
        if raw is None:
            continue
        text = str(raw).strip()
        if text:
            return text
    return None


class ProductionCapacityService:
    """Orchestrates DayCapacityRepository + core.production.capacity helpers."""

    def __init__(
        self,
        *,
        db_path: str | None = None,
        day_capacity_repository: DayCapacityRepository | None = None,
    ) -> None:
        if day_capacity_repository is not None:
            self.day_capacity_repository = day_capacity_repository
            self.db_path = day_capacity_repository.db_path
        else:
            self.db_path = str(db_path or get_settings().plita_db_path)
            self.day_capacity_repository = DayCapacityRepository(db_path=self.db_path)

    def get_capacity_map(self, dates: Sequence[date | str]) -> dict[date, int]:
        overrides = self.day_capacity_repository.list_overrides()
        result: dict[date, int] = {}
        for raw in dates:
            day = _to_date(raw)
            result[day] = int(get_day_capacity(day, overrides))
        return result

    def set_day_capacity(
        self,
        day: date | str,
        max_tracks: int,
        user: Mapping[str, Any] | None = None,
    ) -> None:
        updated_by = _extract_updated_by(user)
        try:
            self.day_capacity_repository.set_override(
                day,
                int(max_tracks),
                updated_by=updated_by,
            )
        except ValueError as exc:
            raise ProductionCapacityError(str(exc)) from exc

    def validate_fill_targets(
        self,
        fill_targets: Sequence[Mapping[str, Any]],
        occupancy: Mapping[str, int] | None = None,
    ) -> None:
        dates: list[date | str] = []
        for target in fill_targets:
            day_raw = target.get("date")
            if day_raw is not None and str(day_raw).strip():
                dates.append(day_raw)  # type: ignore[arg-type]

        capacity_map = self.get_capacity_map(dates)
        day_capacity = {day.isoformat(): max_tracks for day, max_tracks in capacity_map.items()}
        try:
            validate_capacity_fill_targets(
                fill_targets,
                day_capacity,
                occupancy=occupancy,
            )
        except PlanBuildError as exc:
            raise ProductionCapacityError(str(exc)) from exc
