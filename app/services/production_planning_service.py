"""Canonical orchestrator for production plan build/save flows (A10 / WP3).

Единая точка входа для web/bot адаптеров: validate → load → optimize → persist
через ``core.production.planning``. Распределение дорожек — в
``PlanDistributionService`` / ``app.planning.plan_*``; persist — ``PlanRepository`` CRUD;
только legacy-фасад для frozen bot paths.
"""
from __future__ import annotations

import logging
from typing import Any, Literal

from app.core.settings import get_settings
from app.planning.plan_storage import MAX_TRACKS_PER_DAY, create_plan_id
from app.repositories.plan_repository import PlanRepository
from app.services.plan_distribution_service import (
    PlanDistributionService,
    PlanLoadAdapter,
    PlanPersistAdapter,
)
from core.production.dto import (
    LoadConfig,
    LoadResult,
    OptimizeConfig,
    OptimizeResult,
    PersistConfig,
    PlanBuildInput,
)
from core.production.errors import PlanBuildError
from core.plate_order_context import PlateOrderContext
from core.production.planning import load, optimize, persist
from core.rest_matching_service import RestMatchingService

logger = logging.getLogger(__name__)

FilterMethod = Literal["all", "kp"]


class ProductionPlanBuildError(RuntimeError):
    """Доменная ошибка сборки плана (валидное но нестроящееся состояние)."""


def _map_plan_build_error(exc: PlanBuildError) -> ProductionPlanBuildError:
    return ProductionPlanBuildError(str(exc))


class ProductionPlanningService:
    """Orchestrator: собирает вход, вызывает core-pipeline, делегирует persist в repository."""

    def __init__(
        self,
        *,
        plita_db_path: str | None = None,
        pb_db_path: str | None = None,
        plan_repository: PlanRepository | None = None,
    ) -> None:
        settings = get_settings()
        self.plita_db_path = str(plita_db_path or settings.plita_db_path)
        self.pb_db_path = str(pb_db_path or settings.pb_db_path)
        self.plan_repository = plan_repository or PlanRepository(
            db_path=self.plita_db_path
        )
        self.plan_distribution = PlanDistributionService()

    def find_matching_rests(
        self,
        *,
        length_m: float,
        width_mm: int,
        qty_needed: int,
        db_path: str | None = None,
    ) -> list[dict[str, Any]]:
        """Подбор остатков со склада для планирования."""
        return RestMatchingService.find_matching_rests(
            length_m=length_m,
            width_mm=width_mm,
            qty_needed=qty_needed,
            db_path=db_path or self.plita_db_path,
        )

    def run_planning_pipeline(
        self,
        *,
        plan_input: PlanBuildInput | None = None,
        load_result: LoadResult | None = None,
        layout_reinforcement_order: str = "asc",
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> tuple[LoadResult, OptimizeResult]:
        """validate+load → optimize через core; без persist (bot preview / build_plan)."""
        if load_result is None:
            if plan_input is None:
                raise ValueError("Нужен plan_input или load_result.")
            try:
                load_result = load(
                    plan_input,
                    config=LoadConfig(
                        plita_db_path=self.plita_db_path,
                        pb_db_path=self.pb_db_path,
                    ),
                    plan_load=PlanLoadAdapter(self.plan_repository),
                )
            except PlanBuildError as exc:
                raise _map_plan_build_error(exc) from exc

        opt_raw = self._run_optimization_and_split(
            orders_2d=load_result.orders_2d,
            layout_reinforcement_order=layout_reinforcement_order,
            plate_order_ctx=plate_order_ctx,
        )
        if isinstance(opt_raw, tuple):
            all_tracks_list, optimization_result = opt_raw
            opt_result = OptimizeResult(
                all_tracks_list=all_tracks_list,
                optimization_result=optimization_result,
            )
        else:
            opt_result = opt_raw

        return load_result, opt_result

    def build_plan_structure(
        self,
        load_result: LoadResult,
        opt_result: OptimizeResult,
        *,
        start_date: str,
        tracks_count: int,
        layout_reinforcement_order: str = "asc",
    ) -> dict[str, Any]:
        """Собирает структуру плана без commit/persist (bot preview / parity)."""
        if not opt_result.all_tracks_list:
            raise ProductionPlanBuildError("Оптимизация не дала результата.")

        global_occupancy = self.plan_repository.get_global_occupancy()
        plan, stats = self.plan_distribution.build_plan_from_tracks(
            self.plan_repository,
            plan_id=None,
            new_tracks_list=opt_result.all_tracks_list,
            start_date=start_date,
            tracks_per_day=tracks_count,
            plate_lookup_exact=load_result.plate_lookup_exact,
            plate_lookup_by_length=load_result.plate_lookup_by_length,
            orders_2d=load_result.orders_2d,
            optimization_result=opt_result.optimization_result,
            auto_save=False,
            global_occupancy=global_occupancy,
        )
        plan["layout_reinforcement_order"] = layout_reinforcement_order
        return {"plan": plan, "stats": stats}

    def build_plan(
        self,
        *,
        start_date: str,
        tracks_count: int,
        filter_method: FilterMethod,
        selected_kp_ids: list[int] | None = None,
        selected_plate_ids: dict[int, list[int]] | None = None,
        selected_plate_qty: dict[int, dict[int, int]] | None = None,
        active_plan_id: str | None = None,
        plan_name: str | None = None,
        fill_targets: list[dict[str, Any]] | None = None,
        layout_reinforcement_order: str = "asc",
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        """Собирает план по заданным фильтрам через core-pipeline."""
        plan_input = PlanBuildInput(
            start_date=start_date,
            tracks_count=tracks_count,
            filter_method=filter_method,
            selected_kp_ids=tuple(selected_kp_ids or ()),
            selected_plate_ids=selected_plate_ids,
            selected_plate_qty=selected_plate_qty,
            layout_reinforcement_order=layout_reinforcement_order,  # type: ignore[arg-type]
        )

        try:
            load_result, opt_result = self.run_planning_pipeline(
                plan_input=plan_input,
                layout_reinforcement_order=layout_reinforcement_order,
                plate_order_ctx=plate_order_ctx,
            )

            persist_port = PlanPersistAdapter(
                self.plan_repository,
                self.plan_distribution,
            )
            persist_result = persist(
                load_result,
                opt_result,
                PersistConfig(
                    plita_db_path=self.plita_db_path,
                    start_date=start_date,
                    tracks_count=tracks_count,
                    layout_reinforcement_order=layout_reinforcement_order,  # type: ignore[arg-type]
                    active_plan_id=active_plan_id,
                    plan_name=plan_name,
                    fill_targets=tuple(fill_targets or ()),
                    max_tracks_per_day=MAX_TRACKS_PER_DAY,
                ),
                persist_port,
                ensure_unique_plan_id=create_plan_id,
            )
        except PlanBuildError as exc:
            raise _map_plan_build_error(exc) from exc

        return {
            "plan": persist_result.plan,
            "stats": persist_result.stats,
            "summary": persist_result.summary,
        }

    def _run_optimization_and_split(
        self,
        *,
        orders_2d: list[dict[str, Any]],
        layout_reinforcement_order: str = "asc",
        plate_order_ctx: PlateOrderContext | None = None,
    ):
        """Делегирует в core.optimize; оставлен для подмены в тестах."""
        from core.production.dto import LoadResult, OptimizeConfig

        load_result = LoadResult(
            kp_list=[],
            selected_plates=[],
            orders_2d=orders_2d,
            plate_lookup_exact={},
            plate_lookup_by_length={},
        )
        return optimize(
            load_result,
            config=OptimizeConfig(
                pb_db_path=self.pb_db_path,
                layout_reinforcement_order=layout_reinforcement_order,  # type: ignore[arg-type]
                track_top_up_from_following=get_settings().track_top_up_from_following,
            ),
            plate_order_ctx=plate_order_ctx,
        )
