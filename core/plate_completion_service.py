"""Plate completion domain orchestration (A2) — write-off loop and row matching."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence

from core.domain.plate_completion_matching import find_kp_plate_row
from core.domain.plate_completion_types import CompletePlatesResult, UnmovedPlateInfo
from core.kp_db_common import DEFAULT_DB, _connect
from core.kp_db_plates_common import (
    _fetch_kp_plate_row_by_id,
    _purge_zero_qty_plates,
    _record_plate_completion,
)


class PlateCompletionService:
    """Domain service for day completion write-off (orchestration + row lookup)."""

    @staticmethod
    def find_one_row(
        cur: sqlite3.Cursor,
        plate_name: str,
        length_m: float,
        width_m: float,
        load_class: int,
        prefer_kp_id: int,
        *,
        length_dm_raw: str | None = None,
        allow_cross_kp: bool = False,
        plan_ids: Sequence[str] | None = None,
    ) -> tuple | None:
        """Deprecated alias — use :func:`find_kp_plate_row` directly in new code."""
        return find_kp_plate_row(
            cur,
            plate_name,
            length_m,
            width_m,
            load_class,
            prefer_kp_id,
            length_dm_raw=length_dm_raw,
            allow_cross_kp=allow_cross_kp,
            plan_ids=plan_ids,
        )

    @staticmethod
    def complete_plates_on_cursor(
        cur: sqlite3.Cursor,
        kp_id: int,
        plates_to_complete: List[Dict[str, Any]],
        production_day: int,
        *,
        plan_ids: Optional[List[str]] = None,
        allow_cross_kp: bool = False,
        actor: str | None = None,
    ) -> CompletePlatesResult:
        completed_count = 0
        unmoved_plates: list[UnmovedPlateInfo] = []
        completed_date = datetime.now().strftime("%d.%m.%Y")

        for plate in plates_to_complete:
            plate_name = plate.get("plate_name", "")
            qty_remaining = plate.get("qty", 1)
            length_m = plate.get("length_m", 0)
            width_m = plate.get("width_m", 0)
            load_class = plate.get("load_class", 800)
            length_dm_raw = plate.get("length_dm_raw") or ""
            kp_plate_id = plate.get("kp_plate_id")

            if not plate_name:
                continue

            current_plate_name = plate_name
            current_width_m = width_m
            while qty_remaining > 0:
                row = None
                if kp_plate_id:
                    row = _fetch_kp_plate_row_by_id(
                        cur, int(kp_plate_id), production_day, plan_ids
                    )
                if row is None and kp_plate_id:
                    break
                if row is None:
                    row = find_kp_plate_row(
                        cur,
                        current_plate_name,
                        length_m,
                        current_width_m,
                        load_class,
                        kp_id,
                        length_dm_raw=length_dm_raw,
                        allow_cross_kp=allow_cross_kp,
                        plan_ids=plan_ids,
                    )
                if not row:
                    break

                row_id, row_kp_id, row_plate_name, row_width_m, row_qty, row_nomenclature_id = row
                deduct = min(qty_remaining, row_qty)
                _record_plate_completion(
                    cur,
                    row_id=row_id,
                    row_kp_id=row_kp_id,
                    row_plate_name=row_plate_name,
                    length_m=length_m,
                    row_width_m=row_width_m,
                    load_class=load_class,
                    deduct=deduct,
                    completed_date=completed_date,
                    production_day=production_day,
                    row_nomenclature_id=row_nomenclature_id,
                    plan_ids=plan_ids,
                    actor=actor,
                )
                qty_remaining -= deduct
                completed_count += deduct
                current_plate_name = row_plate_name
                current_width_m = row_width_m

                if deduct > 0 and (row_kp_id != kp_id):
                    print(
                        f"[DB] ⚠️ Плита списана из КП #{row_kp_id}: {row_plate_name} (qty={deduct})"
                    )

            if qty_remaining > 0:
                unmoved_plates.append(
                    {
                        "kp_id": int(kp_id),
                        "plate_name": str(current_plate_name or plate_name or ""),
                        "qty": int(qty_remaining),
                        "length_m": float(length_m or 0),
                        "width_m": float(current_width_m or width_m or 0),
                        "load_class": int(load_class or 0),
                    }
                )
                print(
                    f"[DB] ⚠️ Не найдена плита для списания: КП #{kp_id}, {current_plate_name} "
                    f"(width={current_width_m}, осталось qty={qty_remaining})"
                )

        _purge_zero_qty_plates(cur)
        return {"completed_count": completed_count, "unmoved": unmoved_plates}

    @staticmethod
    def move_plates_to_completed(
        kp_id: int,
        plates_to_complete: List[Dict[str, Any]],
        production_day: int,
        db_path: str = DEFAULT_DB,
        plan_ids: Optional[List[str]] = None,
        allow_cross_kp: bool = False,
        *,
        actor: str | None = None,
        return_unmoved: bool = False,
        _external_conn: Optional[sqlite3.Connection] = None,
    ) -> int | tuple[int, list[UnmovedPlateInfo]]:
        own_conn = _external_conn is None
        if own_conn:
            conn = _connect(db_path)
        else:
            conn = _external_conn

        try:
            if own_conn:
                conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.cursor()
            result = PlateCompletionService.complete_plates_on_cursor(
                cur,
                kp_id,
                plates_to_complete,
                production_day,
                plan_ids=plan_ids,
                allow_cross_kp=allow_cross_kp,
                actor=actor,
            )
            completed_count = result["completed_count"]
            unmoved_plates = result["unmoved"]

            if own_conn:
                conn.commit()
            print(
                f"[DB] ✅ Перенесено {completed_count} плит в completed_plates "
                f"(КП #{kp_id}, день {production_day})"
            )
            if return_unmoved:
                return completed_count, unmoved_plates
            return completed_count

        except Exception as e:
            if own_conn:
                print(f"[DB] ❌ Ошибка при переносе плит: {e}")
                conn.rollback()
                if return_unmoved:
                    return 0, []
                return 0
            raise

        finally:
            if own_conn:
                conn.close()
