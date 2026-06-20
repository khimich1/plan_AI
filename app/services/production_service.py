from __future__ import annotations

from datetime import date, datetime
from typing import Any

from app.repositories.kp_repository import KpRepository
from app.repositories.plan_repository import PlanRepository
from app.repositories.work_calendar_repository import WorkCalendarRepository
from app.services.day_view_service import build_day_view_detail
from app.services.optimization_service import OptimizationService
from app.services.production_completion_service import ProductionCompletionService
from app.services.production_planning_service import ProductionPlanningService
from app.planning import plan_manager
from core.plate_order_context import PlateOrderContext
from core.plan_track_removal import TrackRemovalError

MAX_TRACKS_PER_DAY = plan_manager.MAX_TRACKS_PER_DAY

_TRACK_REMOVAL_HTTP_STATUS: dict[str, int] = {
    "plan_not_found": 404,
    "day_not_found": 404,
    "day_already_completed": 409,
    "invalid_track_index": 400,
    "no_plate_identity": 400,
    "incomplete_return": 409,
    "db_return_failed": 500,
    "plan_save_failed": 500,
}


class ProductionTrackRemovalError(Exception):
    """Ошибка удаления дорожки из производственного плана."""

    def __init__(self, message: str, *, status_code: int, code: str | None = None) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.code = code


class ProductionService:
    def __init__(
        self,
        *,
        kp_repository: KpRepository | None = None,
        plan_repository: PlanRepository | None = None,
        calendar_repository: WorkCalendarRepository | None = None,
        optimization_service: OptimizationService | None = None,
        planning_service: ProductionPlanningService | None = None,
        completion_service: ProductionCompletionService | None = None,
    ) -> None:
        self.kp_repository = kp_repository or KpRepository()
        self.plan_repository = plan_repository or PlanRepository()
        self.calendar_repository = calendar_repository or WorkCalendarRepository()
        self.optimization_service = optimization_service or OptimizationService()
        self.planning_service = planning_service or ProductionPlanningService()
        self.completion_service = completion_service or ProductionCompletionService(
            db_path=self.kp_repository.db_path,
            plan_repository=self.plan_repository,
        )

    def list_plans(self) -> dict:
        return self.plan_repository.list_metadata()

    def get_plan(self, plan_id: str) -> dict | None:
        record = self.plan_repository.get(plan_id)
        if not record:
            return None
        return {**record["payload"], "version": record["version"]}

    def activate_plan(self, plan_id: str) -> dict | None:
        if not self.plan_repository.set_active(plan_id):
            return None
        return {"plan_id": plan_id, "active": True}

    def delete_plan(self, plan_id: str) -> dict:
        deleted = self.plan_repository.delete(plan_id)
        return {"plan_id": plan_id, "deleted": deleted}

    def get_day_occupancy(self, exclude_plan_id: str | None = None) -> dict:
        occupancy = self.plan_repository.get_global_occupancy(exclude_plan_id=exclude_plan_id)
        return {
            "occupancy": {str(k): int(v) for k, v in occupancy.items()},
            "max_per_day": int(MAX_TRACKS_PER_DAY),
        }

    def list_kp_candidates(self) -> dict:
        items = self.kp_repository.list_kps_in_production()
        visible = [
            item
            for item in items
            if item.get("in_plan_pct", 0) < 100 and item.get("plates")
        ]
        return {"items": visible, "count": len(visible)}

    def build_plan_from_filters(
        self,
        *,
        start_date: str,
        tracks_count: int,
        filter_method: str,
        selected_kp_ids: list[int] | None = None,
        selected_plate_ids: dict[int, list[int]] | None = None,
        selected_plate_qty: dict[int, dict[int, int]] | None = None,
        active_plan_id: str | None = None,
        plan_name: str | None = None,
        fill_targets: list[dict[str, Any]] | None = None,
        layout_reinforcement_order: str = "asc",
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        return self.planning_service.build_plan(
            start_date=start_date,
            tracks_count=tracks_count,
            filter_method=filter_method,  # type: ignore[arg-type]
            selected_kp_ids=selected_kp_ids,
            selected_plate_ids=selected_plate_ids,
            selected_plate_qty=selected_plate_qty,
            active_plan_id=active_plan_id,
            plan_name=plan_name,
            fill_targets=fill_targets,
            layout_reinforcement_order=layout_reinforcement_order,
            plate_order_ctx=plate_order_ctx,
        )

    def get_calendar(self) -> dict | None:
        return self.plan_repository.get_global_calendar_info()

    def get_day_view(self, target_date: str) -> dict | None:
        return self.plan_repository.get_tracks_for_date(target_date)

    def get_day_view_detailed(self, target_date: str) -> dict | None:
        return build_day_view_detail(
            target_date,
            db_path=self.kp_repository.db_path,
            plan_repository=self.plan_repository,
        )

    def complete_day(
        self,
        *,
        plan_id: str,
        target_date: str,
        rejected_plates: list[dict[str, Any]] | None = None,
        actor: str | None = None,
    ) -> dict:
        completion_result = self.completion_service.complete_day(
            plan_id=plan_id,
            target_date=target_date,
            rejected_plates=rejected_plates,
            actor=actor,
        )
        completed = self.plan_repository.mark_day_completed(plan_id, target_date)
        return {
            "plan_id": plan_id,
            "date": target_date,
            "completed": completed,
            **completion_result,
        }

    def load_candidates_for_plan(self, limit: int = 500) -> list[dict]:
        return self.kp_repository.list_production_candidates(limit=limit)

    def create_plan(
        self,
        *,
        name: str,
        start_date: str,
        tracks_per_day: int,
        all_tracks_list: list,
        plate_lookup_exact: dict | None = None,
        plate_lookup_by_length: dict | None = None,
        orders_2d: list | None = None,
        optimization_result: dict | None = None,
        active_plan_id: str | None = None,
        auto_save: bool = True,
    ) -> dict[str, Any]:
        updated_plan, stats = self.plan_repository.build_plan_from_tracks(
            plan_id=active_plan_id,
            new_tracks_list=all_tracks_list,
            start_date=start_date,
            tracks_per_day=tracks_per_day,
            plate_lookup_exact=plate_lookup_exact or {},
            plate_lookup_by_length=plate_lookup_by_length or {},
            orders_2d=orders_2d or [],
            optimization_result=optimization_result or {},
            auto_save=auto_save,
        )
        updated_plan["name"] = name or updated_plan.get("name") or f"План {updated_plan['id']}"
        if auto_save:
            self.plan_repository.save_plan(updated_plan)
        plan_version = None
        if auto_save:
            saved = self.plan_repository.get(updated_plan["id"])
            plan_version = saved["version"] if saved else None
        result = {"plan": updated_plan, "stats": stats}
        if plan_version is not None:
            result["plan"] = {**updated_plan, "version": plan_version}
        return result

    def get_work_calendar(self) -> dict:
        return self.calendar_repository.load_raw()

    def save_work_calendar(self, payload: dict) -> dict:
        self.calendar_repository.save_raw(payload)
        return self.calendar_repository.load_raw()

    def nth_working_day(self, start_day: str, n: int) -> str:
        parsed = date.fromisoformat(start_day)
        return self.calendar_repository.nth_working_day(parsed, n).isoformat()

    def is_working_day(self, target_date: str) -> bool:
        parsed = datetime.fromisoformat(target_date).date()
        return self.calendar_repository.is_working_day(parsed)

    def remove_track(
        self,
        plan_id: str,
        date: str,
        track_index: int,
        *,
        actor: str | None = None,
    ) -> dict[str, Any]:
        try:
            return self.plan_repository.remove_track_from_plan(
                plan_id,
                date,
                track_index,
                db_path=self.kp_repository.db_path,
                actor=actor,
            )
        except TrackRemovalError as exc:
            status_code = _TRACK_REMOVAL_HTTP_STATUS.get(exc.code or "", 500)
            raise ProductionTrackRemovalError(
                exc.message,
                status_code=status_code,
                code=exc.code,
            ) from exc

