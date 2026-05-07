from __future__ import annotations

from app.core.settings import get_settings
from app.planning import plan_manager


class PlanRepository:
    def __init__(self) -> None:
        self.settings = get_settings()

    def list_metadata(self) -> dict:
        return plan_manager.load_plans_metadata()

    def get_active_plan_id(self) -> str | None:
        return plan_manager.get_active_plan_id()

    def set_active_plan(self, plan_id: str) -> None:
        plan_manager.set_active_plan(plan_id)

    def load_plan(self, plan_id: str) -> dict | None:
        return plan_manager.load_plan(plan_id)

    def save_plan(self, plan_data: dict) -> None:
        plan_manager.save_plan(plan_data)

    def delete_plan(self, plan_id: str) -> bool:
        return plan_manager.delete_plan(plan_id)

    def mark_day_completed(self, plan_id: str, date_key: str) -> bool:
        return plan_manager.mark_day_completed(plan_id, date_key)

    def get_global_occupancy(self, exclude_plan_id: str | None = None) -> dict[str, int]:
        return plan_manager.get_global_day_occupancy(exclude_plan_id=exclude_plan_id)

    def get_tracks_for_date(self, date_key: str) -> dict | None:
        return plan_manager.get_tracks_for_date_from_all_plans(date_key)

    def build_plan_from_tracks(
        self,
        *,
        plan_id: str | None,
        new_tracks_list: list,
        start_date: str,
        tracks_per_day: int,
        plate_lookup_exact: dict | None = None,
        plate_lookup_by_length: dict | None = None,
        orders_2d: list | None = None,
        optimization_result: dict | None = None,
        auto_save: bool = False,
    ) -> tuple[dict, dict]:
        return plan_manager.add_tracks_to_plan(
            plan_id=plan_id,
            new_tracks_list=new_tracks_list,
            start_date=start_date,
            tracks_per_day=tracks_per_day,
            plate_lookup_exact=plate_lookup_exact or {},
            plate_lookup_by_length=plate_lookup_by_length or {},
            orders_2d=orders_2d or [],
            optimization_result=optimization_result or {},
            auto_save=auto_save,
        )

