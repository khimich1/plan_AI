from __future__ import annotations

import logging
import time
from dataclasses import asdict
from datetime import date, datetime, timedelta
from typing import Any, Mapping, Sequence
from zoneinfo import ZoneInfo

from app.repositories.kp_repository import KpRepository
from app.repositories.plan_repository import PlanRepository
from app.repositories.promise_repository import PromiseRepository
from app.repositories.work_calendar_repository import WorkCalendarRepository
from app.services.day_view_service import build_day_view_detail
from app.services.optimization_service import OptimizationService
from app.services.production_capacity_service import (
    ProductionCapacityError,
    ProductionCapacityService,
)
from app.services.production_completion_service import ProductionCompletionService
from app.planning.plan_storage import MAX_TRACKS_PER_DAY
from app.services.production_planning_service import ProductionPlanningService
from app.services.production_substrate_service import (
    ProductionSubstrateError,
    ProductionSubstrateService,
)
from app.services.production_urgent_service import ProductionUrgentService
from core.domain.enums import PlateStatus
from core.execution_terms import parse_execution_terms_to_datetime
from core.plate_order_context import PlateOrderContext
from core.plan_track_removal import TrackRemovalError
from core.production.capacity import FUTURE_HORIZON_DAYS, calculate_capacity_deficit
from core.work_calendar import is_working_day, load_extra_workdays, load_holidays

logger = logging.getLogger(__name__)

_MOSCOW_TZ = ZoneInfo("Europe/Moscow")
_UNPARSED_DEADLINE = date.max


def _in_work_sort_key(item: Mapping[str, Any]) -> tuple:
    parsed = parse_execution_terms_to_datetime(str(item.get("execution_terms") or ""))
    kp_id = int(item["kp_id"])
    if parsed is None:
        return (1, _UNPARSED_DEADLINE, kp_id)
    return (0, parsed.date(), kp_id)


def _with_bucket(plates: Sequence[Mapping[str, Any]], bucket: str) -> list[dict[str, Any]]:
    return [{**dict(plate), "bucket": bucket} for plate in plates]


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


class ProductionAnalyzeBadRequest(ValueError):
    """Невалидные даты / параметры анализа подложек (HTTP 400)."""


class ProductionAnalyzeEmptyBacklog(RuntimeError):
    """Нет плит «в производстве» для анализа (HTTP 422)."""


def _to_date(value: date | str) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


def _iso_weeks_covering(from_date: date, to_date: date) -> list[date]:
    start = from_date - timedelta(days=from_date.weekday())
    last = to_date - timedelta(days=to_date.weekday())
    weeks: list[date] = []
    cursor = start
    while cursor <= last:
        weeks.append(cursor)
        cursor += timedelta(days=7)
    return weeks


def _candidate_promise_meta(
    allocs: Sequence[Mapping[str, Any]] | None,
) -> dict[str, Any] | None:
    if not allocs:
        return None
    overdue = [row for row in allocs if row["status"] == "overdue"]
    primary = min(overdue or allocs, key=lambda row: row["week_start"])
    return {
        "promised_date": primary["promised_date"],
        "week_start": primary["week_start"],
        "status": "overdue" if overdue else "active",
        "tracks": int(primary["tracks_total"]),
    }


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
        promise_repository: PromiseRepository | None = None,
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
        self.promise_repository = promise_repository or PromiseRepository(
            db_path=self.kp_repository.db_path
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
        from app.services.sgp_service import SgpService

        SgpService(db_path=self.kp_repository.db_path).clear_plan_links(plan_id)
        deleted = self.plan_repository.delete(plan_id)
        return {"plan_id": plan_id, "deleted": deleted}

    def get_day_occupancy(self, exclude_plan_id: str | None = None) -> dict:
        occupancy = self.plan_repository.get_global_occupancy(exclude_plan_id=exclude_plan_id)
        occupancy_map = {str(k): int(v) for k, v in occupancy.items()}

        capacity_dates: list[date] = []
        for key in occupancy_map:
            try:
                capacity_dates.append(date.fromisoformat(str(key)))
            except ValueError:
                continue

        # Empty occupancy: short horizon so FE still gets per-day max without a huge map.
        if not capacity_dates:
            today = datetime.now(_MOSCOW_TZ).date()
            capacity_dates = [
                today + timedelta(days=offset)
                for offset in range(FUTURE_HORIZON_DAYS + 1)
            ]

        capacity_service = ProductionCapacityService(db_path=self.kp_repository.db_path)
        capacity_map = capacity_service.get_capacity_map(capacity_dates)
        max_by_day = {
            day.isoformat(): int(max_tracks)
            for day, max_tracks in capacity_map.items()
        }

        return {
            "occupancy": occupancy_map,
            "max_per_day": int(MAX_TRACKS_PER_DAY),
            "max_by_day": max_by_day,
        }

    def list_kp_candidates(
        self,
        *,
        scope: str = "plan",
        from_date: date | None = None,
        to_date: date | None = None,
    ) -> dict:
        items = self.kp_repository.list_kps_in_production()
        if scope == "in_work":
            payload = self._candidates_in_work(items)
        else:
            payload = self._candidates_plan(items)
        return self._attach_promise_meta(
            payload, from_date=from_date, to_date=to_date
        )

    def _attach_promise_meta(
        self,
        payload: Mapping[str, Any],
        *,
        from_date: date | None,
        to_date: date | None,
    ) -> dict[str, Any]:
        week_starts = None
        if from_date is not None and to_date is not None:
            week_starts = _iso_weeks_covering(from_date, to_date)
        allocs = self.promise_repository.list_wizard_promise_allocs(
            week_starts=week_starts,
        )
        by_kp: dict[int, list[dict[str, Any]]] = {}
        by_week: dict[date, list[dict[str, Any]]] = {}
        for alloc in allocs:
            by_kp.setdefault(int(alloc["kp_id"]), []).append(alloc)
            by_week.setdefault(alloc["week_start"], []).append(alloc)

        items = []
        for item in payload["items"]:
            meta = _candidate_promise_meta(by_kp.get(int(item["kp_id"])))
            items.append({**dict(item), "promise": meta})
        promised_weeks = [
            {
                "week_start": week,
                "items": [
                    {
                        "kp_id": int(row["kp_id"]),
                        "promised_date": row["promised_date"],
                        "tracks": int(row["tracks"]),
                        "status": row["status"],
                    }
                    for row in rows
                ],
            }
            for week, rows in sorted(by_week.items())
        ]
        return {**dict(payload), "items": items, "promised_weeks": promised_weeks}

    def _candidates_plan(self, items: Sequence[Mapping[str, Any]]) -> dict:
        visible: list[dict[str, Any]] = []
        for item in items:
            if float(item.get("in_plan_pct") or 0) >= 100:
                continue
            plates = list(item.get("plates") or [])
            if not plates:
                continue
            visible.append({**dict(item), "plates": _with_bucket(plates, "awaiting_plan")})
        return {"items": visible, "count": len(visible)}

    def _candidates_in_work(self, items: Sequence[Mapping[str, Any]]) -> dict:
        kp_ids = [int(item["kp_id"]) for item in items]
        planned_by_kp = self.kp_repository.fetch_plates_in_statuses(
            kp_ids,
            (PlateStatus.IN_PLAN.value,),
        )
        visible: list[dict[str, Any]] = []
        for item in items:
            remaining = int(item.get("remaining_qty") or 0)
            in_plan = int(item.get("in_plan_qty") or 0)
            if remaining + in_plan <= 0:
                continue
            awaiting = _with_bucket(list(item.get("plates") or []), "awaiting_plan")
            planned = _with_bucket(
                planned_by_kp.get(int(item["kp_id"]), []),
                "in_plan",
            )
            visible.append({**dict(item), "plates": awaiting + planned})
        visible.sort(key=_in_work_sort_key)
        return {"items": visible, "count": len(visible)}

    def analyze_substrates(
        self,
        *,
        fill_targets: Sequence[Mapping[str, Any]],
        deadline_until: date | str,
        user: Mapping[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Urgent positions + substrate recommendations + capacity deficit."""
        _ = user  # reserved for audit / capacity overrides attribution
        deadline = _to_date(deadline_until)
        targets = [dict(item) for item in fill_targets]

        if not targets:
            raise ProductionAnalyzeBadRequest("fill_targets пуст.")

        fill_dates: list[date] = []
        for item in targets:
            day_raw = item.get("date")
            try:
                fill_dates.append(_to_date(str(day_raw)))
            except ValueError as exc:
                raise ProductionAnalyzeBadRequest(
                    f"Неверный формат даты в fill_targets: {day_raw}"
                ) from exc

        first_fill = min(fill_dates)
        if deadline < first_fill:
            raise ProductionAnalyzeBadRequest(
                "deadline_until раньше первой даты fill_targets."
            )

        occupancy = {
            str(k): int(v)
            for k, v in self.plan_repository.get_global_occupancy().items()
        }
        capacity_service = ProductionCapacityService(db_path=self.kp_repository.db_path)
        try:
            capacity_service.validate_fill_targets(targets, occupancy=occupancy)
        except ProductionCapacityError as exc:
            raise ProductionAnalyzeBadRequest(str(exc)) from exc

        kps = self.kp_repository.list_kps_in_production()
        length_by_plate_id: dict[int, float] = {}
        orders_count = 0
        for kp in kps:
            for plate in kp.get("plates") or []:
                plate_id = int(plate["id"])
                length_by_plate_id[plate_id] = float(plate.get("length_m") or 0.0)
                orders_count += 1

        if orders_count == 0:
            raise ProductionAnalyzeEmptyBacklog(
                "Нет плит «в производстве» для анализа."
            )

        urgent_service = ProductionUrgentService(kp_repository=self.kp_repository)
        urgent = urgent_service.list_urgent_positions(deadline_until=deadline)

        substrate_service = ProductionSubstrateService(
            kp_repository=self.kp_repository
        )
        started = time.perf_counter()
        optimization_status = "ok"
        error_message: str | None = None
        try:
            substrates = substrate_service.find_substrate_recommendations(
                urgent_plate_ids=[int(p.plate_id) for p in urgent],
                deadline_until=deadline,
                first_fill_target_date=first_fill,
            )
        except ProductionSubstrateError as exc:
            logger.exception(
                "[analyze_substrates] Ошибка анализа подложек: %s", exc
            )
            substrates = []
            optimization_status = "error"
            error_message = str(exc) or "Ошибка анализа подложек"
        analysis_duration_ms = int((time.perf_counter() - started) * 1000)

        urgent_length_m = 0.0
        for position in urgent:
            length_m = length_by_plate_id.get(int(position.plate_id), 0.0)
            urgent_length_m += length_m * int(position.qty_remaining)

        today = datetime.now(_MOSCOW_TZ).date()
        min_fill = min(fill_dates)
        max_fill = max(fill_dates)
        capacity_dates: list[date] = []
        cursor = today
        while cursor < min_fill:
            capacity_dates.append(cursor)
            cursor += timedelta(days=1)
        capacity_dates.extend(fill_dates)
        cursor = max_fill + timedelta(days=1)
        horizon_end = max_fill + timedelta(days=FUTURE_HORIZON_DAYS)
        while cursor <= horizon_end:
            capacity_dates.append(cursor)
            cursor += timedelta(days=1)

        capacity_map = capacity_service.get_capacity_map(capacity_dates)
        day_capacity = {
            day.isoformat(): int(max_tracks)
            for day, max_tracks in capacity_map.items()
        }
        completed_dates = self._completed_plan_dates()
        holidays = load_holidays()
        extra_workdays = load_extra_workdays()

        deficit = calculate_capacity_deficit(
            urgent_length_m,
            targets,
            day_capacity,
            occupancy=occupancy,
            completed_dates=completed_dates,
            today=today,
            is_workday=lambda d: is_working_day(
                d, holidays=holidays, extra_workdays=extra_workdays
            ),
        )

        return {
            "urgent_positions": [asdict(p) for p in urgent],
            "substrate_recommendations": [asdict(s) for s in substrates],
            "capacity_deficit": asdict(deficit) if deficit is not None else None,
            "analysis_meta": {
                "orders_count": orders_count,
                "analysis_duration_ms": analysis_duration_ms,
                "optimization_status": optimization_status,
                "error_message": error_message,
            },
        }

    def _completed_plan_dates(self) -> set[str]:
        completed: set[str] = set()
        for plan in self.plan_repository.list_all_plans():
            for date_key, day_data in (plan.get("days") or {}).items():
                if day_data.get("completed"):
                    completed.add(str(date_key))
        return completed

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
        sgp_reservations: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        from app.services.sgp_service import SgpService
        from core import kp_db_plates
        import sqlite3

        qty_for_optimize = selected_plate_qty
        from_sgp_rows: list[dict[str, Any]] = []
        if sgp_reservations:
            sgp = SgpService(db_path=self.kp_repository.db_path)
            qty_for_optimize = sgp.reduce_selected_qty_for_reservations(
                selected_plate_qty=selected_plate_qty,
                sgp_reservations=sgp_reservations,
            )
            from_sgp_rows = sgp.build_from_sgp_rows(sgp_reservations)

        result = self.planning_service.build_plan(
            start_date=start_date,
            tracks_count=tracks_count,
            filter_method=filter_method,  # type: ignore[arg-type]
            selected_kp_ids=selected_kp_ids,
            selected_plate_ids=selected_plate_ids,
            selected_plate_qty=qty_for_optimize,
            active_plan_id=active_plan_id,
            plan_name=plan_name,
            fill_targets=fill_targets,
            layout_reinforcement_order=layout_reinforcement_order,
            plate_order_ctx=plate_order_ctx,
        )
        if sgp_reservations:
            plan = result.get("plan") or {}
            plan_id = plan.get("id")
            sgp = SgpService(db_path=self.kp_repository.db_path)
            conn = sqlite3.connect(self.kp_repository.db_path)
            try:
                conn.execute("PRAGMA foreign_keys = ON")
                cur = conn.cursor()
                reserved_total = 0
                for item in sgp_reservations:
                    reserved_total += sgp.reserve_on_conn(
                        cur,
                        conn,
                        sgp_id=int(item["sgp_id"]),
                        target_kp_id=int(item["target_kp_id"]),
                        qty=int(item["qty"]),
                        plan_id=plan_id,
                    )
                plan["sgp_reservations"] = list(sgp_reservations)
                plan["from_sgp_qty"] = reserved_total
                plan["from_sgp"] = from_sgp_rows
                if plan_id:
                    self.plan_repository.save_plan(plan)
                conn.commit()
                result["sgp_reserved_qty"] = reserved_total
            except Exception:
                # D4: plan already persisted by build_plan — compensate, do not
                # leave a plan with sgp_reservations without an actual reserve.
                conn.rollback()
                if plan_id:
                    try:
                        kp_db_plates.return_plan_plates_to_production(
                            str(plan_id), self.kp_repository.db_path
                        )
                        self.plan_repository.delete(str(plan_id))
                    except Exception:
                        logger.exception(
                            "SGP build compensate failed for plan %s", plan_id
                        )
                raise
            finally:
                conn.close()
        return result

    def get_calendar(self) -> dict | None:
        return self.planning_service.plan_distribution.get_global_calendar_info(
            self.plan_repository
        )

    def get_day_view(self, target_date: str) -> dict | None:
        return self.planning_service.plan_distribution.get_tracks_for_date(
            self.plan_repository,
            target_date,
        )

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
        expected_version: int | None = None,
    ) -> dict:
        # D4: KP write-off + mark_day_completed — одна tx внутри completion_service.
        completion_result = self.completion_service.complete_day(
            plan_id=plan_id,
            target_date=target_date,
            rejected_plates=rejected_plates,
            actor=actor,
            expected_version=expected_version,
        )
        return {
            "plan_id": plan_id,
            "date": target_date,
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
        updated_plan, stats = self.planning_service.plan_distribution.build_plan_from_tracks(
            self.plan_repository,
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
        expected_version: int | None = None,
    ) -> dict[str, Any]:
        try:
            return self.planning_service.plan_distribution.remove_track_from_plan(
                self.plan_repository,
                plan_id,
                date,
                track_index,
                db_path=self.kp_repository.db_path,
                actor=actor,
                expected_version=expected_version,
            )
        except TrackRemovalError as exc:
            status_code = _TRACK_REMOVAL_HTTP_STATUS.get(exc.code or "", 500)
            raise ProductionTrackRemovalError(
                exc.message,
                status_code=status_code,
                code=exc.code,
            ) from exc

