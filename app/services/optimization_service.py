from __future__ import annotations

import copy
import warnings
from contextlib import contextmanager
from typing import Any, Iterator

import core.optimization as legacy_optimization

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from core.optimization import result_contract as _opt_contract
from core.optimization import optimize_with_cascading_longitudinal_cuts
from core.plate_order_context import PlateOrderContext


class OptimizationService:
    def optimize(
        self,
        order: PlateOrder,
        *,
        orders_2d: list[dict[str, Any]] | None = None,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> OptimizationContext:
        # Для совместимости с новым web-пайплайном допускаем явный orders_2d:
        # там важны дополнительные поля (kp_id/plate_name) для последующего
        # корректного commit в БД.
        source_orders_2d = orders_2d if orders_2d is not None else order.to_orders_2d()

        def _run_optimization() -> dict[str, Any]:
            if not source_orders_2d:
                return _opt_contract.opt_error(
                    _opt_contract.ERROR_EMPTY_ORDERS_2D,
                    "Пустой список для 2D оптимизации (PlateOrder / orders_2d).",
                )
            return optimize_with_cascading_longitudinal_cuts(orders_2d=source_orders_2d)

        if plate_order_ctx is not None:
            plate_order_ctx.hydrate_from_order(order)
            with plate_order_ctx.bound():
                result = _run_optimization()
        else:
            result = _run_optimization()
        all_loads = (
            sorted({int(float(item.get("load_code", 8))) for item in source_orders_2d})
            if source_orders_2d
            else [8]
        )
        if _opt_contract.is_optimization_success(result):
            result["loads_in_group"] = all_loads
        plan_by_load = {"all": result} if _opt_contract.is_optimization_success(result) else {}
        load_map = {load: ["all"] for load in all_loads} if _opt_contract.is_optimization_success(result) else {}
        return OptimizationContext(
            order=order,
            optimization_result=result,
            plan_by_load=plan_by_load,
            load_to_reinforcement_map=load_map,
        )

    @contextmanager
    def bound_plate_order_context(
        self,
        plate_ctx: PlateOrderContext,
        optimization_context: OptimizationContext,
    ) -> Iterator[PlateOrderContext]:
        """Привязать OPT-снимок к ``PlateOrderContext`` (предпочтительный путь A1)."""
        plate_ctx.load_optimization_snapshot(
            optimization_result=optimization_context.optimization_result,
            plan_by_load=optimization_context.plan_by_load,
            load_to_reinforcement_map=optimization_context.load_to_reinforcement_map,
        )
        with plate_ctx.bound():
            yield plate_ctx

    @contextmanager
    def legacy_runtime(self, context: OptimizationContext) -> Iterator[OptimizationContext]:
        """Deprecated: используйте ``bound_plate_order_context`` + явный ``PlateOrderContext``."""
        warnings.warn(
            "OptimizationService.legacy_runtime() is deprecated; "
            "use bound_plate_order_context(plate_ctx, optimization_context) instead.",
            DeprecationWarning,
            stacklevel=2,
        )
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
