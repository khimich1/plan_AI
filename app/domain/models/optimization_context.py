from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from app.domain.models.plate_order import PlateOrder
from core.optimization import result_contract as _opt_contract


@dataclass
class OptimizationContext:
    order: PlateOrder
    optimization_result: dict[str, Any] = field(default_factory=dict)
    plan_by_load: dict[str, Any] = field(default_factory=dict)
    load_to_reinforcement_map: dict[Any, list[Any]] = field(default_factory=dict)

    @property
    def optimization_success(self) -> bool:
        return _opt_contract.is_optimization_success(self.optimization_result)

    @property
    def optimization_status(self) -> str:
        r = self.optimization_result
        if not r:
            return "error"
        s = r.get(_opt_contract.OPT_STATUS_KEY)
        if s in ("ok", "error", "partial"):
            return s
        return "ok" if self.optimization_success else "error"

    @property
    def optimization_error_code(self) -> str | None:
        return _opt_contract.optimization_error_code(self.optimization_result)

    @property
    def optimization_error_message(self) -> str | None:
        return _opt_contract.optimization_error_message(self.optimization_result)

    @property
    def total_plates(self) -> int:
        return int(self.optimization_result.get("total_plates", 0) or 0)

    @property
    def total_cost(self) -> float:
        return float(self.optimization_result.get("total_cost", 0.0) or 0.0)

