from __future__ import annotations

from datetime import date, datetime
from typing import Any

from app.repositories.kp_repository import KpRepository
from app.repositories.plan_repository import PlanRepository
from app.repositories.work_calendar_repository import WorkCalendarRepository
from app.services.optimization_service import OptimizationService
from bot.handlers import plan_manager


class ProductionService:
    def __init__(self) -> None:
        self.kp_repository = KpRepository()
        self.plan_repository = PlanRepository()
        self.calendar_repository = WorkCalendarRepository()
        self.optimization_service = OptimizationService()

    def list_plans(self) -> dict:
        return self.plan_repository.list_metadata()

    def get_plan(self, plan_id: str) -> dict | None:
        return self.plan_repository.load_plan(plan_id)

    def activate_plan(self, plan_id: str) -> dict:
        self.plan_repository.set_active_plan(plan_id)
        return {"plan_id": plan_id, "active": True}

    def get_calendar(self) -> dict | None:
        return plan_manager.get_global_calendar_info()

    def get_day_view(self, target_date: str) -> dict | None:
        return self.plan_repository.get_tracks_for_date(target_date)

    def complete_day(self, *, plan_id: str, target_date: str) -> dict:
        completed = self.plan_repository.mark_day_completed(plan_id, target_date)
        return {"plan_id": plan_id, "date": target_date, "completed": completed}

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
        return {"plan": updated_plan, "stats": stats}

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

