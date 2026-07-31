#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared plate persistence helpers (A1 slice)."""

from __future__ import annotations

import sqlite3
from typing import List, Optional

from app.domain.enums import PlateStatus, PlateTransitionReason
from core.kp_db_audit import audit_append
from core.kp_db_common import DEFAULT_DB, _connect


def _normalize_plate_name(name: str) -> str:
    """
    Нормализует имя плиты для совпадений (DEPRECATED — оставлен для обратной
    совместимости со старым кодом). Используйте :func:`core.plate_name.canonical`.
    """
    from core import plate_name as _pn

    return _pn.canonical(name)


def _fetch_kp_plate_row_by_id(
    cur: sqlite3.Cursor,
    kp_plate_id: int,
    production_day: int,
    plan_ids: Optional[List[str]] = None,
) -> tuple | None:
    """Lookup kp_plates row by primary key (plan/day constrained)."""
    try:
        if plan_ids:
            placeholders = ",".join("?" * len(plan_ids))
            cur.execute(
                f"""
                SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id
                FROM kp_plates
                WHERE id = ?
                  AND plan_id IN ({placeholders})
                  AND day_number = ?
                  AND status IN ('в плане', 'в производстве')
                  AND qty > 0
                LIMIT 1
                """,
                (int(kp_plate_id), *plan_ids, int(production_day)),
            )
        else:
            cur.execute(
                """
                SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id
                FROM kp_plates
                WHERE id = ?
                  AND day_number = ?
                  AND status IN ('в плане', 'в производстве')
                  AND qty > 0
                LIMIT 1
                """,
                (int(kp_plate_id), int(production_day)),
            )
        return cur.fetchone()
    except Exception:
        return None


def _deduct_kp_plate_qty(cur: sqlite3.Cursor, row_id: int, deduct: int) -> None:
    cur.execute("UPDATE kp_plates SET qty = qty - ? WHERE id = ?", (deduct, row_id))


def _insert_completed_plate(
    cur: sqlite3.Cursor,
    *,
    row_kp_id: int | None,
    row_plate_name: str,
    length_m: float,
    row_width_m: float,
    load_class: int,
    deduct: int,
    completed_date: str,
    production_day: int,
    row_nomenclature_id,
    plan_id: str | None = None,
) -> None:
    cur.execute(
        """
        INSERT INTO completed_plates (
            kp_id, plate_name, length_m, width_m, load_class,
            qty, completed_date, production_day, nomenclature_id, plan_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            row_kp_id,
            row_plate_name,
            length_m,
            row_width_m,
            load_class,
            deduct,
            completed_date,
            production_day,
            row_nomenclature_id,
            plan_id,
        ),
    )


def _record_plate_completion(
    cur: sqlite3.Cursor,
    *,
    row_id: int,
    row_kp_id: int,
    row_plate_name: str,
    length_m: float,
    row_width_m: float,
    load_class: int,
    deduct: int,
    completed_date: str,
    production_day: int,
    row_nomenclature_id,
    plan_ids: Optional[List[str]],
    actor: str | None,
) -> None:
    plan_id = plan_ids[0] if plan_ids else None
    _deduct_kp_plate_qty(cur, row_id, deduct)
    _insert_completed_plate(
        cur,
        row_kp_id=row_kp_id,
        row_plate_name=row_plate_name,
        length_m=length_m,
        row_width_m=row_width_m,
        load_class=load_class,
        deduct=deduct,
        completed_date=completed_date,
        production_day=production_day,
        row_nomenclature_id=row_nomenclature_id,
        plan_id=plan_id,
    )
    audit_append(
        cur,
        plate_id=row_id,
        kp_id=row_kp_id,
        plate_name=row_plate_name,
        plan_id=plan_id,
        day_number=production_day,
        from_status=PlateStatus.IN_PLAN.value,
        to_status=PlateStatus.ON_SGP.value,
        qty=deduct,
        reason=PlateTransitionReason.SGP_SEND.value,
        actor=actor,
    )


def _purge_zero_qty_plates(cur: sqlite3.Cursor) -> None:
    cur.execute("DELETE FROM kp_plates WHERE qty <= 0")


def insert_kp_plate_remainder_row(
    cur: sqlite3.Cursor,
    *,
    source_plate_id: int,
    remainder_qty: int,
    status: str = "в производстве",
    plan_id: str | None = None,
    day_number: int | None = None,
) -> int:
    """
    DRY split INSERT: new kp_plates row copied from ``source_plate_id`` (Q1/A13).

    Preserves ``nomenclature_id``, ``length_dm_raw``, weights, and price fields.
    """
    cur.execute(
        """
        INSERT INTO kp_plates (
            kp_id, position_number, plate_name, length_m, width_m, load_class,
            qty, unit_weight, total_weight, discounted_price, status, plan_id, day_number,
            concrete_grade, nomenclature_id, length_dm_raw
        )
        SELECT
            kp_id, position_number, plate_name, length_m, width_m, load_class,
            ?, unit_weight, total_weight, discounted_price, ?, ?, ?,
            concrete_grade, nomenclature_id, length_dm_raw
        FROM kp_plates WHERE id = ?
        """,
        (remainder_qty, status, plan_id, day_number, source_plate_id),
    )
    return int(cur.lastrowid)
