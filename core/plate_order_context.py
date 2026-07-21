#!/usr/bin/env python3

# -*- coding: utf-8 -*-

"""

Единый контекст заказа плит + OPT-состояния на один HTTP-запрос / прогон бота (A1-001).



``bound()`` вкладывает ``plate_mutable_runtime_scope`` и ``optimization_context_scope``,

чтобы legacy-доступ через ``get_plate_mutable_runtime()`` и OPT_*-прокси продолжал работать

без смены call sites (strangler).

"""



from __future__ import annotations



import asyncio

import copy

from collections.abc import Callable, Iterator

from contextlib import contextmanager

from dataclasses import dataclass

from typing import Any



from core.domain.plate_order import PlateOrder as CorePlateOrder

from core.domain.plate_order import _try_fill_plate_nomenclature_cache, coerce_core_plate_order

from core.optimization.context import (

    new_optimization_context_state,

    optimization_context_scope,

)

from core.optimization.result_contract import is_optimization_success

from core.plate_runtime_state import (

    PlateMutableRuntime,

    new_plate_mutable_runtime_empty,

    plate_mutable_runtime_scope,

)

from core.runtime import NomenclatureCacheFiller





@dataclass

class PlateOrderContext:

    """SSOT мутабельного заказа и OPT_* на время запроса или явного прогона."""



    plates: PlateMutableRuntime

    optimization: dict[str, Any]



    @classmethod

    def fresh_empty(cls) -> PlateOrderContext:

        """Пустой заказ и свежее OPT-состояние (как middleware S1)."""

        return cls(

            plates=new_plate_mutable_runtime_empty(),

            optimization=new_optimization_context_state(),

        )



    @contextmanager

    def bound(self) -> Iterator[PlateOrderContext]:

        """Привязать plates + optimization к текущей asyncio-задаче / потоку."""

        with plate_mutable_runtime_scope(self.plates):

            with optimization_context_scope(self.optimization):

                yield self



    def hydrate_from_order(

        self,

        plate_order: Any,

        *,

        fill_nomenclature_cache: NomenclatureCacheFiller | None = None,

    ) -> None:

        """Заполнить ``self.plates`` из ``PlateOrder`` (замена ``apply_to_globals`` в A1-002)."""

        core_order = coerce_core_plate_order(plate_order)



        rt = self.plates

        rt.plate_load_details.clear()

        rt.plate_load_details.update(core_order.plate_load_details)

        rt.plate_length_dm_raw.clear()

        rt.plate_length_dm_raw.update(core_order.plate_length_dm_raw)

        rt.plate_nomenclature_cache.clear()

        rt.plates_1_2 = list(core_order.plates_1_2)

        rt.plates_1_5_to_1_2 = list(core_order.plates_1_5_to_1_2)

        rt.plates_1_0 = list(core_order.plates_1_0)

        rt.plates_1_08 = list(core_order.plates_1_08)

        rt.plates_0_46 = list(core_order.plates_0_46)

        rt.plates_0_32 = list(core_order.plates_0_32)

        rt.plates_0_72 = list(core_order.plates_0_72)

        rt.plates_0_70 = list(core_order.plates_0_70)

        rt.plates_0_86 = list(core_order.plates_0_86)

        rt.plates_0_74 = list(core_order.plates_0_74)

        rt.plates_0_88 = list(core_order.plates_0_88)

        rt.plates_0_48 = list(core_order.plates_0_48)

        rt.plates_0_50 = list(core_order.plates_0_50)

        rt.plates_0_34 = list(core_order.plates_0_34)

        rt.plate_exact_widths.clear()

        rt.plate_exact_widths.update(core_order.plate_exact_widths)

        rt.longitudinal_cuts = core_order.longitudinal_cuts

        rt.length_trims = core_order.length_trims

        rt.unused_strips_0_3_m_total = core_order.unused_strips_0_3_m_total

        rt.scrap_strips_0_2_m_total = core_order.scrap_strips_0_2_m_total

        rt.usable_strips_0_74_m_total = core_order.usable_strips_0_74_m_total

        rt.usable_strips_0_88_m_total = core_order.usable_strips_0_88_m_total

        rt.usable_strips_0_48_m_total = core_order.usable_strips_0_48_m_total

        rt.usable_strips_0_50_m_total = core_order.usable_strips_0_50_m_total

        rt.usable_strips_0_34_m_total = core_order.usable_strips_0_34_m_total

        rt.scrap_strips_0_12_m_total = core_order.scrap_strips_0_12_m_total

        rt.waste_area_m2 = core_order.waste_area_m2



        with self.bound():

            _try_fill_plate_nomenclature_cache(fill_nomenclature_cache)



    def snapshot_core_order(self) -> CorePlateOrder:

        """Снимок канонического заказа из ``self.plates``."""

        return CorePlateOrder.from_runtime(self.plates)



    def load_optimization_snapshot(

        self,

        *,

        optimization_result: dict | None = None,

        plan_by_load: dict | None = None,

        load_to_reinforcement_map: dict | None = None,

    ) -> None:

        """Записать OPT_*-снимок в ``self.optimization`` (для worker с ``bound()``)."""

        opt = self.optimization

        if optimization_result is not None:

            opt["opt_cascading_plan"].clear()

            opt["opt_cascading_plan"].update(copy.deepcopy(optimization_result))

        if plan_by_load is not None:

            opt["opt_cascading_plan_by_load"].clear()

            opt["opt_cascading_plan_by_load"].update(copy.deepcopy(plan_by_load))

        if load_to_reinforcement_map is not None:

            opt["load_to_reinforcement_map"].clear()

            opt["load_to_reinforcement_map"].update(copy.deepcopy(load_to_reinforcement_map))



    def load_production_snapshot(

        self,

        orders_2d: list,

        optimization_result: dict,

    ) -> None:

        """Заказ + результат оптимизации для ``visualize_plan`` (day/archive/bot)."""

        self.hydrate_from_order(CorePlateOrder.from_orders_2d(orders_2d))



        if orders_2d:

            load_codes = sorted({p.get("load_code", 8) for p in orders_2d})

            load_map = {code: ["all"] for code in load_codes}

        else:

            load_map = {8: ["all"]}



        plan_by_load = (

            {"all": optimization_result}

            if is_optimization_success(optimization_result)

            else {}

        )

        self.load_optimization_snapshot(

            optimization_result=optimization_result,

            plan_by_load=plan_by_load,

            load_to_reinforcement_map=load_map,

        )





async def run_in_order_context(

    ctx: PlateOrderContext,

    fn: Callable[..., Any],

    *args: Any,

    **kwargs: Any,

) -> Any:

    """Выполнить sync ``fn`` в worker-потоке с ``ctx.bound()`` (asyncio.to_thread)."""



    def _worker() -> Any:

        with ctx.bound():

            return fn(*args, **kwargs)



    return await asyncio.to_thread(_worker)





__all__ = [

    "PlateOrderContext",

    "run_in_order_context",

]

