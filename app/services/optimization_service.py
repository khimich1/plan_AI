from __future__ import annotations

import copy
from contextlib import contextmanager
from typing import Any, Iterator

import core.optimization as legacy_optimization

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from core.optimization import optimize_with_cascading_longitudinal_cuts


class OptimizationService:
    def optimize(
        self,
        order: PlateOrder,
        *,
        orders_2d: list[dict[str, Any]] | None = None,
    ) -> OptimizationContext:
        # Для совместимости с новым web-пайплайном допускаем явный orders_2d:
        # там важны дополнительные поля (kp_id/plate_name) для последующего
        # корректного commit в БД.
        source_orders_2d = orders_2d if orders_2d is not None else order.to_orders_2d()
        result = (
            optimize_with_cascading_longitudinal_cuts(orders_2d=source_orders_2d)
            if source_orders_2d
            else {}
        )
        all_loads = (
            sorted({int(float(item.get("load_code", 8))) for item in source_orders_2d})
            if source_orders_2d
            else [8]
        )
        if result:
            result["loads_in_group"] = all_loads
        plan_by_load = {"all": result} if result else {}
        load_map = {load: ["all"] for load in all_loads} if result else {}
        return OptimizationContext(
            order=order,
            optimization_result=result,
            plan_by_load=plan_by_load,
            load_to_reinforcement_map=load_map,
        )

    @contextmanager
    def legacy_runtime(self, context: OptimizationContext) -> Iterator[OptimizationContext]:
        snapshot = {
            "OPT_CASCADING_PLAN": copy.deepcopy(legacy_optimization.OPT_CASCADING_PLAN),
            "OPT_CASCADING_PLAN_BY_LOAD": copy.deepcopy(legacy_optimization.OPT_CASCADING_PLAN_BY_LOAD),
            "LOAD_TO_REINFORCEMENT_MAP": copy.deepcopy(legacy_optimization.LOAD_TO_REINFORCEMENT_MAP),
        }
        try:
            legacy_optimization.OPT_CASCADING_PLAN = copy.deepcopy(context.optimization_result)
            legacy_optimization.OPT_CASCADING_PLAN_BY_LOAD = copy.deepcopy(context.plan_by_load)
            legacy_optimization.LOAD_TO_REINFORCEMENT_MAP = copy.deepcopy(context.load_to_reinforcement_map)
            yield context
        finally:
            legacy_optimization.OPT_CASCADING_PLAN = snapshot["OPT_CASCADING_PLAN"]
            legacy_optimization.OPT_CASCADING_PLAN_BY_LOAD = snapshot["OPT_CASCADING_PLAN_BY_LOAD"]
            legacy_optimization.LOAD_TO_REINFORCEMENT_MAP = snapshot["LOAD_TO_REINFORCEMENT_MAP"]

