from __future__ import annotations

import json
import logging
import sqlite3
from datetime import datetime, timedelta
from typing import Any

from app.core.settings import get_settings
from app.planning import plan_manager
from app.planning.plan_aggregation import (
    _iter_plan_tracks_for_date,
    _merge_plate_lookups,
    get_all_tracks_from_plan,
)
from app.planning.plan_storage import MAX_TRACKS_PER_DAY, count_day_tracks
from app.repositories.plan_errors import PlanVersionConflict
from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.plan_track_removal import TrackRemovalError, collect_plate_returns_from_track
from core.serialization import strip_plate_audit_from_plan

logger = logging.getLogger(__name__)


class PlanRepository:
    def __init__(self, db_path: str | None = None) -> None:
        self.settings = get_settings()
        self.db_path = db_path or str(self.settings.plita_db_path)

    def _connect(self) -> sqlite3.Connection:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    @staticmethod
    def _require_plan_id(payload: dict[str, Any]) -> str:
        plan_id = payload.get("id")
        if not plan_id or not isinstance(plan_id, str):
            raise ValueError("plan payload must contain a non-empty string 'id'")
        return plan_id

    @staticmethod
    def _serialize_payload(payload: dict[str, Any]) -> str:
        clean = strip_plate_audit_from_plan(payload)
        return json.dumps(clean, ensure_ascii=False)

    @staticmethod
    def _deserialize_payload(payload_json: str) -> dict[str, Any]:
        data = json.loads(payload_json)
        if not isinstance(data, dict):
            raise ValueError("stored plan payload must be a JSON object")
        return data

    @staticmethod
    def _payload_to_metadata_entry(
        payload: dict[str, Any],
        *,
        version: int | None = None,
    ) -> dict[str, Any]:
        plan_id = payload["id"]
        days = payload.get("days", {}) or {}
        total_tracks = sum(count_day_tracks(day) for day in days.values())
        entry = {
            "id": plan_id,
            "name": payload.get("name", f"План {plan_id}"),
            "created_at": payload.get(
                "created_at",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
            "start_date": payload.get("start_date", ""),
            "total_days": len(days),
            "tracks_count": payload.get("tracks_count", 5),
            "total_tracks": total_tracks,
        }
        if version is not None:
            entry["version"] = version
        return entry

    def _list_all_plans(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT payload_json
                FROM production_plans
                ORDER BY created_at ASC, id ASC
                """
            ).fetchall()
        return [self._deserialize_payload(row["payload_json"]) for row in rows]

    def get(self, plan_id: str) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT payload_json, version
                FROM production_plans
                WHERE id = ?
                """,
                (plan_id,),
            ).fetchone()
        if row is None:
            return None
        return {
            "payload": self._deserialize_payload(row["payload_json"]),
            "version": int(row["version"]),
        }

    def create(self, payload: dict[str, Any]) -> dict[str, Any]:
        plan_id = self._require_plan_id(payload)
        payload_json = self._serialize_payload(payload)
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO production_plans (id, payload_json, version, is_active)
                VALUES (?, ?, 1, 0)
                """,
                (plan_id, payload_json),
            )
            conn.commit()
        return {"payload": self._deserialize_payload(payload_json), "version": 1}

    def save(self, payload: dict[str, Any], expected_version: int) -> dict[str, Any]:
        plan_id = self._require_plan_id(payload)
        payload_json = self._serialize_payload(payload)
        with self._connect() as conn:
            cursor = conn.execute(
                """
                UPDATE production_plans
                SET payload_json = ?,
                    version = version + 1,
                    updated_at = datetime('now')
                WHERE id = ? AND version = ?
                """,
                (payload_json, plan_id, expected_version),
            )
            if cursor.rowcount == 0:
                raise PlanVersionConflict(plan_id, expected_version)
            row = conn.execute(
                """
                SELECT payload_json, version
                FROM production_plans
                WHERE id = ?
                """,
                (plan_id,),
            ).fetchone()
            conn.commit()
        if row is None:
            raise PlanVersionConflict(plan_id, expected_version)
        return {
            "payload": self._deserialize_payload(row["payload_json"]),
            "version": int(row["version"]),
        }

    def delete(self, plan_id: str) -> bool:
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM production_plans WHERE id = ?",
                (plan_id,),
            )
            conn.commit()
        return cursor.rowcount > 0

    def _list_metadata_from_sqlite(self) -> dict[str, Any]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, payload_json, is_active, version
                FROM production_plans
                ORDER BY created_at ASC, id ASC
                """
            ).fetchall()
        plans: list[dict[str, Any]] = []
        active_plan_id: str | None = None
        for row in rows:
            payload = self._deserialize_payload(row["payload_json"])
            plans.append(
                self._payload_to_metadata_entry(
                    payload,
                    version=int(row["version"]),
                )
            )
            if int(row["is_active"]) == 1:
                active_plan_id = row["id"]
        return {"plans": plans, "active_plan_id": active_plan_id}

    def list_metadata(self) -> dict:
        return self._list_metadata_from_sqlite()

    def get_active(self) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT payload_json, version
                FROM production_plans
                WHERE is_active = 1
                LIMIT 1
                """
            ).fetchone()
        if row is None:
            return None
        return {
            "payload": self._deserialize_payload(row["payload_json"]),
            "version": int(row["version"]),
        }

    def set_active(self, plan_id: str) -> bool:
        with self._connect() as conn:
            exists = conn.execute(
                "SELECT 1 FROM production_plans WHERE id = ?",
                (plan_id,),
            ).fetchone()
            if exists is None:
                return False
            conn.execute("UPDATE production_plans SET is_active = 0")
            conn.execute(
                "UPDATE production_plans SET is_active = 1, updated_at = datetime('now') WHERE id = ?",
                (plan_id,),
            )
            conn.commit()
        return True

    def get_active_plan_id(self) -> str | None:
        return self._list_metadata_from_sqlite().get("active_plan_id")

    def set_active_plan(self, plan_id: str) -> bool:
        return self.set_active(plan_id)

    def load_plan(self, plan_id: str) -> dict | None:
        record = self.get(plan_id)
        return record["payload"] if record else None

    def save_plan(
        self,
        plan_data: dict,
        *,
        expected_version: int | None = None,
    ) -> dict[str, Any]:
        plan_id = self._require_plan_id(plan_data)
        existing = self.get(plan_id)
        if existing is None:
            return self.create(plan_data)
        version = expected_version if expected_version is not None else existing["version"]
        return self.save(plan_data, expected_version=version)

    def delete_plan(self, plan_id: str) -> bool:
        return self.delete(plan_id)

    def mark_day_completed(
        self,
        plan_id: str,
        date_key: str,
        *,
        expected_version: int | None = None,
    ) -> bool:
        record = self.get(plan_id)
        if not record:
            return False

        plan = record["payload"]
        if date_key not in plan.get("days", {}):
            return False

        plan["days"][date_key]["completed"] = True
        if "completed_days" not in plan:
            plan["completed_days"] = []

        day_number = plan["days"][date_key].get("day_number", 1)
        if day_number not in plan["completed_days"]:
            plan["completed_days"].append(day_number)
            plan["completed_days"].sort()

        version = expected_version if expected_version is not None else record["version"]
        self.save(plan, expected_version=version)
        return True

    def get_global_occupancy(self, exclude_plan_id: str | None = None) -> dict[str, int]:
        occupancy: dict[str, int] = {}
        for plan in self._list_all_plans():
            plan_id = plan.get("id")
            if plan_id == exclude_plan_id:
                continue
            for date_key, day_data in plan.get("days", {}).items():
                tracks_count = count_day_tracks(day_data)
                occupancy[date_key] = occupancy.get(date_key, 0) + tracks_count
        return occupancy

    def get_tracks_for_date(self, date_key: str) -> dict | None:
        plans = self._list_all_plans()
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

    def get_global_calendar_info(self) -> dict | None:
        metadata = self._list_metadata_from_sqlite()
        plans_meta = metadata.get("plans", [])
        if not plans_meta:
            logger.warning("[GLOBAL_CALENDAR] Нет сохранённых планов")
            return None

        all_dates_data: dict[str, dict[str, Any]] = {}
        earliest_date: datetime | None = None
        latest_date: datetime | None = None
        total_tracks_count = 0

        for plan in self._list_all_plans():
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

        for day_offset in range(total_days):
            current_date = earliest_date + timedelta(days=day_offset)
            date_key = current_date.strftime("%Y-%m-%d")
            day_number = day_offset + 1
            date_data = all_dates_data.get(date_key, {"occupied": 0, "completed": False})
            days_info[date_key] = {
                "occupied": date_data["occupied"],
                "max": MAX_TRACKS_PER_DAY,
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

    def get_all_plans_gantt_data(self) -> dict | None:
        plans = self._list_all_plans()
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
        existing_plan = None
        existing_version: int | None = None
        if plan_id:
            record = self.get(plan_id)
            if record:
                existing_plan = record["payload"]
                existing_version = record["version"]

        plan, stats = plan_manager.add_tracks_to_plan(
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
                self.create(plan)
            else:
                if existing_version is None:
                    record = self.get(plan["id"])
                    existing_version = record["version"] if record else 1
                self.save(plan, expected_version=existing_version)

        return plan, stats

    def remove_track_from_plan(
        self,
        plan_id: str,
        date_key: str,
        track_index: int,
        *,
        db_path: str,
        actor: str | None = None,
        expected_version: int | None = None,
    ) -> dict[str, Any]:
        from core import kp_db

        record = self.get(plan_id)
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

        conn = sqlite3.connect(db_path)
        plates_returned = 0
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            db_result = kp_db.return_plate_rows_for_plan(
                plan_id,
                id_qty,
                db_path,
                actor=actor,
                legacy_identity_qty=legacy_identity_qty or None,
                _external_conn=conn,
            )
            plates_returned = int(db_result.get("plates_returned") or 0)
            db_warnings = db_result.get("warnings") or []
            if db_warnings or plates_returned < expected_count:
                conn.rollback()
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
            conn.commit()
        except TrackRemovalError:
            raise
        except Exception as exc:
            conn.rollback()
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
        finally:
            conn.close()

        tracks.pop(track_index)
        day["saved_tracks_count"] = len(tracks)
        saved_tracks_count = day["saved_tracks_count"]

        if not tracks:
            del plan["days"][date_key]
            saved_tracks_count = 0

        version = expected_version if expected_version is not None else stored_version
        self.save(plan, expected_version=version)

        return {
            "plan_id": plan_id,
            "date": date_key,
            "track_index": track_index,
            "plates_returned": plates_returned,
            "saved_tracks_count": saved_tracks_count,
        }
