from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from app.domain.models.plate_order import PlateOrder


@dataclass
class OptimizationContext:
    order: PlateOrder
    optimization_result: dict[str, Any] = field(default_factory=dict)
    plan_by_load: dict[str, Any] = field(default_factory=dict)
    load_to_reinforcement_map: dict[Any, list[Any]] = field(default_factory=dict)

    @property
    def total_plates(self) -> int:
        return int(self.optimization_result.get("total_plates", 0) or 0)

    @property
    def total_cost(self) -> float:
        return float(self.optimization_result.get("total_cost", 0.0) or 0.0)

