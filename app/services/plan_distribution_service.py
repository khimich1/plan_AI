"""Orchestration for plan track distribution and multi-plan aggregation (A2 / P3 WP2)."""
from __future__ import annotations

import logging
from datetime import datetime, timedelta
from typing import Any

from app.planning.plan_aggregation import (
    _iter_plan_tracks_for_date,
    _merge_plate_lookups,
    get_all_tracks_from_plan,
)
from app.planning.plan_distribution import add_tracks_to_plan
from app.planning.plan_storage import MAX_TRACKS_PER_DAY, count_day_tracks
from app.repositories.plan_repository import PlanRepository
from app.services.production_capacity_service import ProductionCapacityService
from core.plan_track_removal import TrackRemovalError, collect_plate_returns_from_track
from core.production.capacity import clamp_day_max
from core.production.dto import FilterMethod
from core.production.ports import PlanLoadPort

logger = logging.getLogger(__name__)


class PlanDistributionService:
    """Business orchestration: track distribution, aggregation, track removal."""

    def build_plan_from_tracks(
        self,
        repo: PlanRepository,
        *,
        plan_id: str | None,
        new_tracks_list: list,
        start_date: str,
        tracks_per_day: int,
        plate_lookup_exact: dict | None = None,
        plate_lookup_by_length: dict | None = None,
        orders_2d: list | None = None,
        optimization_result: dict | None = None,
        plan_name: str | None = None,
        global_occupancy: dict[str, int] | None = None,
        precomputed_tracks_by_day: dict[str, list] | None = None,
        auto_save: bool = False,
    ) -> tuple[dict, dict]:
        existing_plan = None
        existing_version: int | None = None
        if plan_id:
            record = repo.get(plan_id)
            if record:
                existing_plan = record["payload"]
                existing_version = record["version"]

        plan, stats = add_tracks_to_plan(
            plan_id=plan_id,
            new_tracks_list=new_tracks_list,
            start_date=start_date,
            tracks_per_day=tracks_per_day,
            plate_lookup_exact=plate_lookup_exact or {},
            plate_lookup_by_length=plate_lookup_by_length or {},
            orders_2d=orders_2d or [],
            optimization_result=optimization_result or {},
            plan_name=plan_name,
            auto_save=False,
            global_occupancy=global_occupancy,
            precomputed_tracks_by_day=precomputed_tracks_by_day,
            existing_plan=existing_plan,
        )

        if auto_save:
            if stats.get("is_new_plan"):
                repo.create(plan)
            else:
                if existing_version is None:
                    record = repo.get(plan["id"])
                    existing_version = record["version"] if record else 1
                repo.save(plan, expected_version=existing_version)

        return plan, stats

    def get_tracks_for_date(
        self,
        repo: PlanRepository,
        date_key: str,
    ) -> dict | None:
        plans = repo.list_all_plans()
        if not plans:
            logger.warning("[MULTI_PLAN] Нет сохранённых планов для даты %s", date_key)
            return None

        all_tracks_for_date: list = []
        combined_plate_lookup_exact: dict = {}
        combined_plate_lookup_by_length: dict = {}
        combined_orders_2d: list = []
        last_optimization_result: dict = {}
        source_plan_ids: list[str] = []
        plans_with_date = 0

        for plan in plans:
            plan_id = plan.get("id")
            if date_key not in plan.get("days", {}):
                continue

            plans_with_date += 1
            source_plan_ids.append(str(plan_id))

            day_tracks = _iter_plan_tracks_for_date(plan, date_key)
            all_tracks_for_date.extend(day_tracks)

            _merge_plate_lookups(
                combined_plate_lookup_exact,
                combined_plate_lookup_by_length,
                plan,
                dedup_mode="kp_plate",
            )

            combined_orders_2d.extend(plan.get("orders_2d", []))
            plan_opt_result = plan.get("optimization_result", {})
            if plan_opt_result:
                last_optimization_result = plan_opt_result

        if plans_with_date == 0:
            logger.warning("[MULTI_PLAN] Дата %s не найдена ни в одном плане", date_key)
            return None

        return {
            "tracks": all_tracks_for_date,
            "plate_lookup_exact": combined_plate_lookup_exact,
            "plate_lookup_by_length": combined_plate_lookup_by_length,
            "orders_2d": combined_orders_2d,
            "optimization_result": last_optimization_result,
            "plans_count": plans_with_date,
            "source_plans": source_plan_ids,
        }

    def get_global_calendar_info(self, repo: PlanRepository) -> dict | None:
        metadata = repo.list_metadata()
        plans_meta = metadata.get("plans", [])
        if not plans_meta:
            logger.warning("[GLOBAL_CALENDAR] Нет сохранённых планов")
            return None

        all_dates_data: dict[str, dict[str, Any]] = {}
        earliest_date: datetime | None = None
        latest_date: datetime | None = None
        total_tracks_count = 0

        for plan in repo.list_all_plans():
            for date_key, day_data in plan.get("days", {}).items():
                try:
                    day_dt = datetime.strptime(date_key, "%Y-%m-%d")
                    if earliest_date is None or day_dt < earliest_date:
                        earliest_date = day_dt
                    if latest_date is None or day_dt > latest_date:
                        latest_date = day_dt
                except ValueError:
                    logger.warning("[GLOBAL_CALENDAR] Неверный формат даты: %s", date_key)
                    continue

                tracks_count = count_day_tracks(day_data)
                is_completed = day_data.get("completed", False)

                if date_key not in all_dates_data:
                    all_dates_data[date_key] = {"occupied": 0, "completed": False}

                all_dates_data[date_key]["occupied"] += tracks_count
                if is_completed:
                    all_dates_data[date_key]["completed"] = True

                total_tracks_count += tracks_count

        if earliest_date is None or latest_date is None:
            logger.warning("[GLOBAL_CALENDAR] Не удалось определить диапазон дат")
            return None

        total_days = (latest_date - earliest_date).days + 1
        days_info: dict[str, dict[str, Any]] = {}
        completed_days: list[int] = []

        capacity_day_keys = [
            (earliest_date + timedelta(days=offset)).date()
            for offset in range(total_days)
        ]
        try:
            capacity_svc = ProductionCapacityService()
            capacity_map = capacity_svc.get_capacity_map(capacity_day_keys)
        except Exception:
            logger.exception(
                "[GLOBAL_CALENDAR] Failed to load day capacity; using default max"
            )
            capacity_map = {}

        for day_offset in range(total_days):
            current_date = earliest_date + timedelta(days=day_offset)
            date_key = current_date.strftime("%Y-%m-%d")
            day_number = day_offset + 1
            date_data = all_dates_data.get(date_key, {"occupied": 0, "completed": False})
            day_max = capacity_map.get(current_date.date(), MAX_TRACKS_PER_DAY)
            days_info[date_key] = {
                "occupied": date_data["occupied"],
                "max": clamp_day_max(int(day_max)),
                "completed": date_data["completed"],
                "day_number": day_number,
            }
            if date_data["completed"]:
                completed_days.append(day_number)

        return {
            "start_date": earliest_date.strftime("%Y-%m-%d"),
            "total_days": total_days,
            "days_info": days_info,
            "completed_days": completed_days,
            "plans_count": len(plans_meta),
            "tracks_count": total_tracks_count,
        }

    def get_all_plans_gantt_data(self, repo: PlanRepository) -> dict | None:
        plans = repo.list_all_plans()
        if not plans:
            logger.warning("[GANTT] Нет сохранённых планов для диаграммы")
            return None

        all_tracks_combined: list = []
        combined_plate_lookup_exact: dict = {}
        combined_plate_lookup_by_length: dict = {}
        earliest_date: datetime | None = None
        latest_date: datetime | None = None
        unique_dates: set[str] = set()
        plans_loaded = 0

        for plan in plans:
            plans_loaded += 1
            all_tracks_combined.extend(get_all_tracks_from_plan(plan))
            _merge_plate_lookups(
                combined_plate_lookup_exact,
                combined_plate_lookup_by_length,
                plan,
                dedup_mode="entry",
            )

            plan_start = plan.get("start_date")
            if plan_start:
                try:
                    start_dt = datetime.strptime(plan_start, "%Y-%m-%d")
                    if earliest_date is None or start_dt < earliest_date:
                        earliest_date = start_dt
                except ValueError:
                    logger.warning(
                        "[GANTT] Неверный формат даты начала в плане %s: %s",
                        plan.get("id"),
                        plan_start,
                    )

            for date_key in plan.get("days", {}).keys():
                unique_dates.add(date_key)
                try:
                    day_dt = datetime.strptime(date_key, "%Y-%m-%d")
                    if latest_date is None or day_dt > latest_date:
                        latest_date = day_dt
                except ValueError:
                    logger.warning("[GANTT] Неверный формат даты дня: %s", date_key)

        if plans_loaded == 0 or not all_tracks_combined:
            return None

        if earliest_date is None:
            earliest_date = datetime.now()
        if latest_date is None:
            latest_date = datetime.now()

        return {
            "all_tracks": all_tracks_combined,
            "plate_lookup_exact": combined_plate_lookup_exact,
            "plate_lookup_by_length": combined_plate_lookup_by_length,
            "earliest_start_date": earliest_date,
            "latest_end_date": latest_date,
            "plans_count": plans_loaded,
            "total_days": len(unique_dates),
        }

    def remove_track_from_plan(
        self,
        repo: PlanRepository,
        plan_id: str,
        date_key: str,
        track_index: int,
        *,
        db_path: str,
        actor: str | None = None,
        expected_version: int | None = None,
    ) -> dict[str, Any]:
        record = repo.get(plan_id)
        if not record:
            raise TrackRemovalError(
                f"План {plan_id!r} не найден",
                code="plan_not_found",
            )

        plan = record["payload"]
        stored_version = record["version"]

        day = plan.get("days", {}).get(date_key)
        if day is None:
            raise TrackRemovalError(
                f"День {date_key!r} не найден в плане {plan_id!r}",
                code="day_not_found",
            )

        if day.get("completed"):
            raise TrackRemovalError(
                f"День {date_key!r} уже завершён — удаление дорожки невозможно",
                code="day_already_completed",
            )

        tracks = day.get("tracks") or []
        if track_index < 0 or track_index >= len(tracks):
            raise TrackRemovalError(
                f"Недопустимый track_index={track_index} (дорожек в дне: {len(tracks)})",
                code="invalid_track_index",
            )

        track = tracks[track_index]
        id_qty, legacy_identity_qty = collect_plate_returns_from_track(track)

        if not id_qty and not legacy_identity_qty:
            raise TrackRemovalError(
                "В дорожке не найдено kp_plate_id и legacy-идентичностей — "
                "удаление дорожки невозможно",
                code="no_plate_identity",
            )

        expected_count = sum(id_qty.values()) + sum(legacy_identity_qty.values())

        plates_returned = 0
        try:
            with repo.kp_plates_write_transaction(db_path) as conn:
                db_result = repo.return_plate_rows_for_plan_on_connection(
                    conn,
                    plan_id=plan_id,
                    id_qty=id_qty,
                    legacy_identity_qty=legacy_identity_qty or None,
                    actor=actor,
                    db_path=db_path,
                )
                plates_returned = int(db_result.get("plates_returned") or 0)
                db_warnings = db_result.get("warnings") or []
                if db_warnings or plates_returned < expected_count:
                    detail = (
                        f"ожидалось вернуть {expected_count} плит(ы), "
                        f"фактически {plates_returned}"
                    )
                    if db_warnings:
                        detail = f"{detail}; предупреждения: {'; '.join(db_warnings)}"
                    raise TrackRemovalError(
                        f"Неполный возврат плит в производство: {detail}",
                        code="incomplete_return",
                    )
        except TrackRemovalError:
            raise
        except Exception as exc:
            logger.exception(
                "[REMOVE_TRACK] Ошибка возврата плит plan_id=%s date=%s track_index=%s",
                plan_id,
                date_key,
                track_index,
            )
            raise TrackRemovalError(
                f"Не удалось вернуть плиты в производство: {exc}",
                code="db_return_failed",
            ) from exc

        tracks.pop(track_index)
        day["saved_tracks_count"] = len(tracks)
        saved_tracks_count = day["saved_tracks_count"]

        if not tracks:
            del plan["days"][date_key]
            saved_tracks_count = 0

        version = expected_version if expected_version is not None else stored_version
        repo.save(plan, expected_version=version)

        return {
            "plan_id": plan_id,
            "date": date_key,
            "track_index": track_index,
            "plates_returned": plates_returned,
            "saved_tracks_count": saved_tracks_count,
        }


class PlanLoadAdapter:
    """``PlanLoadPort`` surface for ``core.production.planning.load``."""

    def __init__(self, repo: PlanRepository) -> None:
        self._repo = repo

    def fetch_kps_in_production(
        self,
        *,
        filter_method: FilterMethod,
        selected_kp_ids: list[int],
    ) -> list[tuple[int, str | None, str | None]]:
        return self._repo.fetch_kps_in_production(
            filter_method=filter_method,
            selected_kp_ids=selected_kp_ids,
        )

    def fetch_plates_in_production_for_kp(
        self,
        *,
        kp_id: int,
        plate_ids: list[int] | None,
    ) -> list[tuple[Any, ...]]:
        return self._repo.fetch_plates_in_production_for_kp(
            kp_id=kp_id,
            plate_ids=plate_ids,
        )


class PlanPersistAdapter:
    """``PlanPersistPort`` surface for ``core.production.planning.persist``."""

    def __init__(
        self,
        repo: PlanRepository,
        distribution: PlanDistributionService | None = None,
    ) -> None:
        self._repo = repo
        self._distribution = distribution or PlanDistributionService()

    def get_global_occupancy(
        self, *, exclude_plan_id: str | None = None
    ) -> dict[str, int]:
        return self._repo.get_global_occupancy(exclude_plan_id=exclude_plan_id)

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
        plan_name: str | None = None,
        global_occupancy: dict[str, int] | None = None,
        precomputed_tracks_by_day: dict[str, list] | None = None,
        auto_save: bool = False,
    ) -> tuple[dict, dict]:
        return self._distribution.build_plan_from_tracks(
            self._repo,
            plan_id=plan_id,
            new_tracks_list=new_tracks_list,
            start_date=start_date,
            tracks_per_day=tracks_per_day,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            orders_2d=orders_2d,
            optimization_result=optimization_result,
            plan_name=plan_name,
            global_occupancy=global_occupancy,
            precomputed_tracks_by_day=precomputed_tracks_by_day,
            auto_save=auto_save,
        )

    def get(self, plan_id: str) -> dict[str, Any] | None:
        return self._repo.get(plan_id)

    def create(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self._repo.create(payload)

    def save(self, payload: dict[str, Any], expected_version: int) -> dict[str, Any]:
        return self._repo.save(payload, expected_version=expected_version)

    def set_active(self, plan_id: str) -> None:
        self._repo.set_active(plan_id)
