"""Data contracts for the production planning pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Literal

FilterMethod = Literal["all", "kp"]
LayoutReinforcementOrder = Literal["asc", "desc"]

# Status string for kp_plates rows eligible for planning (mirrors app.domain.enums.PlateStatus).
PLATE_STATUS_IN_PRODUCTION = "в производстве"


@dataclass(frozen=True, slots=True)
class PlanBuildInput:
    """User-facing parameters for building a production plan."""

    start_date: str
    tracks_count: int
    filter_method: FilterMethod
    selected_kp_ids: tuple[int, ...] | None = None
    selected_plate_ids: dict[int, list[int]] | None = None
    selected_plate_qty: dict[int, dict[int, int]] | None = None
    layout_reinforcement_order: LayoutReinforcementOrder = "asc"


@dataclass(frozen=True, slots=True)
class LoadConfig:
    """Database paths and options for the load phase."""

    plita_db_path: str
    pb_db_path: str


@dataclass(frozen=True, slots=True)
class OptimizeConfig:
    """Options for the optimize phase."""

    pb_db_path: str
    layout_reinforcement_order: LayoutReinforcementOrder = "asc"
    track_top_up_from_following: bool = False


@dataclass(slots=True)
class KpEntry:
    kp_id: int
    date: datetime
    customer: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "kp_id": self.kp_id,
            "date": self.date,
            "customer": self.customer,
        }


@dataclass(slots=True)
class LoadResult:
    """Output of validate + load: data ready for optimization."""

    kp_list: list[dict[str, Any]]
    selected_plates: list[dict[str, Any]]
    orders_2d: list[dict[str, Any]]
    plate_lookup_exact: dict[tuple[float, int], list[dict[str, Any]]]
    plate_lookup_by_length: dict[float, list[dict[str, Any]]]


@dataclass(slots=True)
class OptimizeResult:
    """Output of the optimize phase."""

    all_tracks_list: list[dict[str, Any]]
    optimization_result: dict[str, Any] = field(default_factory=dict)


# Mirrors app.planning.plan_storage.MAX_TRACKS_PER_DAY (core must not import app).
DEFAULT_MAX_TRACKS_PER_DAY = 5


@dataclass(frozen=True, slots=True)
class PersistConfig:
    """Parameters for the persist phase (calendar, plan metadata, fill mode)."""

    plita_db_path: str
    start_date: str
    tracks_count: int
    layout_reinforcement_order: LayoutReinforcementOrder = "asc"
    active_plan_id: str | None = None
    plan_name: str | None = None
    fill_targets: tuple[dict[str, Any], ...] | None = None
    max_tracks_per_day: int = DEFAULT_MAX_TRACKS_PER_DAY
    # Per-day max (≤ hard cap). When set, free = day_max − occupied uses this map.
    day_capacity: tuple[tuple[str, int], ...] | None = None


@dataclass(slots=True)
class PersistResult:
    """Output of validate → load → optimize → persist."""

    plan: dict[str, Any]
    stats: dict[str, Any]
    summary: dict[str, Any]
