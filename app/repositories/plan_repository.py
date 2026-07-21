from __future__ import annotations

import json
import logging
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import Any

from app.core.settings import get_settings
from app.planning.plan_storage import count_day_tracks
from app.repositories.plan_errors import PlanVersionConflict
from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.production.dto import FilterMethod, PLATE_STATUS_IN_PRODUCTION
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

    def list_all_plans(self) -> list[dict[str, Any]]:
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

    def delete_all_plans(self) -> int:
        """Delete all rows from ``production_plans`` in a single transaction."""
        with self._connect() as conn:
            cursor = conn.execute("DELETE FROM production_plans")
            deleted = int(cursor.rowcount)
            conn.commit()
        return deleted

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
        for plan in self.list_all_plans():
            plan_id = plan.get("id")
            if plan_id == exclude_plan_id:
                continue
            for date_key, day_data in plan.get("days", {}).items():
                tracks_count = count_day_tracks(day_data)
                occupancy[date_key] = occupancy.get(date_key, 0) + tracks_count
        return occupancy

    def fetch_kps_in_production(
        self,
        *,
        filter_method: FilterMethod,
        selected_kp_ids: list[int],
    ) -> list[tuple[int, str | None, str | None]]:
        """KP rows (kp_id, execution_terms, customer_name) with status «в работе»."""
        with self._connect() as conn:
            if filter_method == "kp":
                placeholders = ",".join("?" * len(selected_kp_ids))
                rows = conn.execute(
                    f"""
                    SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                    FROM KP_offers kp
                    JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                    WHERE kp.kp_id IN ({placeholders})
                      AND meta.status = 'в работе'
                    """,
                    tuple(selected_kp_ids),
                ).fetchall()
            else:
                rows = conn.execute(
                    """
                    SELECT kp.kp_id, kp.execution_terms, kp.customer_name
                    FROM KP_offers kp
                    JOIN kp_meta meta ON kp.kp_id = meta.kp_id
                    WHERE meta.status = 'в работе'
                    """
                ).fetchall()
        return [(int(r[0]), r[1], r[2]) for r in rows]

    def fetch_plates_in_production_for_kp(
        self,
        *,
        kp_id: int,
        plate_ids: list[int] | None,
    ) -> list[tuple[Any, ...]]:
        """Plate rows for one KP in «в производстве» status."""
        with self._connect() as conn:
            if plate_ids:
                placeholders = ",".join("?" * len(plate_ids))
                rows = conn.execute(
                    f"""
                    SELECT id, plate_name, length_m, width_m, load_class, qty, length_dm_raw,
                           COALESCE(concrete_grade, '') AS concrete_grade
                    FROM kp_plates
                    WHERE kp_id = ? AND status = ?
                      AND id IN ({placeholders})
                    ORDER BY position_number, id
                    """,
                    (kp_id, PLATE_STATUS_IN_PRODUCTION) + tuple(plate_ids),
                ).fetchall()
            else:
                rows = conn.execute(
                    """
                    SELECT id, plate_name, length_m, width_m, load_class, qty, length_dm_raw,
                           COALESCE(concrete_grade, '') AS concrete_grade
                    FROM kp_plates
                    WHERE kp_id = ? AND status = ?
                    ORDER BY position_number, id
                    """,
                    (kp_id, PLATE_STATUS_IN_PRODUCTION),
                ).fetchall()
        return [tuple(row) for row in rows]

    @contextmanager
    def kp_plates_write_transaction(self, db_path: str | None = None):
        """Transactional scope for kp_plates mutations (caller validates before exit)."""
        path = db_path or self.db_path
        conn = _connect(path)
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def return_plate_rows_for_plan_on_connection(
        self,
        conn: sqlite3.Connection,
        *,
        plan_id: str,
        id_qty: Any,
        legacy_identity_qty: Any | None = None,
        actor: str | None = None,
        db_path: str | None = None,
    ) -> dict[str, Any]:
        from core import kp_db

        path = db_path or self.db_path
        return kp_db.return_plate_rows_for_plan(
            plan_id,
            id_qty,
            path,
            actor=actor,
            legacy_identity_qty=legacy_identity_qty or None,
            _external_conn=conn,
        )
