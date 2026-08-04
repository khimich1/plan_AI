#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared SQL helpers for shipment reservations (SHIP-2xx).

Резерв — вычислимый: ``completed_plates.qty − Σ shipment_items open-рейсов``.
Единый источник для propose, pre-flight complete и SGP-гарда unlink/relink.
"""

from __future__ import annotations

import sqlite3
from typing import Optional

from core.domain.enums import ShipmentItemType, ShipmentStatus


def allocated_qty_for_completed_plate(
    cur: sqlite3.Cursor,
    completed_plate_id: int,
    *,
    exclude_shipment_id: Optional[int] = None,
) -> int:
    """Штуки строки СГП, зарезервированные open-рейсами (``in_work``)."""
    sql = """
        SELECT COALESCE(SUM(si.qty), 0)
        FROM shipment_items si
        JOIN shipments s ON s.id = si.shipment_id
        WHERE si.completed_plate_id = ?
          AND si.item_type = ?
          AND s.status = ?
    """
    params: list = [
        int(completed_plate_id),
        ShipmentItemType.PLATE.value,
        ShipmentStatus.IN_WORK.value,
    ]
    if exclude_shipment_id is not None:
        sql += " AND s.id != ?"
        params.append(int(exclude_shipment_id))
    cur.execute(sql, params)
    return int(cur.fetchone()[0] or 0)


def available_qty(
    cur: sqlite3.Cursor,
    completed_plate_id: int,
    *,
    exclude_shipment_id: Optional[int] = None,
) -> int:
    """Свободный остаток строки СГП: qty минус резерв open-рейсов."""
    cur.execute(
        "SELECT qty FROM completed_plates WHERE id = ?",
        (int(completed_plate_id),),
    )
    row = cur.fetchone()
    if row is None:
        return 0
    return int(row[0] or 0) - allocated_qty_for_completed_plate(
        cur, completed_plate_id, exclude_shipment_id=exclude_shipment_id
    )


def open_reservation_for_completed_plate(
    cur: sqlite3.Cursor,
    completed_plate_id: int,
) -> Optional[tuple[int, str, int]]:
    """Первый open-рейс, зарезервировавший строку СГП: (shipment_id, date, qty)."""
    cur.execute(
        """
        SELECT s.id, s.shipment_date, SUM(si.qty)
        FROM shipment_items si
        JOIN shipments s ON s.id = si.shipment_id
        WHERE si.completed_plate_id = ?
          AND si.item_type = ?
          AND s.status = ?
        GROUP BY s.id, s.shipment_date
        ORDER BY s.id
        LIMIT 1
        """,
        (
            int(completed_plate_id),
            ShipmentItemType.PLATE.value,
            ShipmentStatus.IN_WORK.value,
        ),
    )
    row = cur.fetchone()
    if row is None:
        return None
    return int(row[0]), str(row[1]), int(row[2] or 0)


def shipped_qty_for_kp(cur: sqlite3.Cursor, kp_id: int) -> int:
    """Σ отгруженных plate-шт по КП из done-рейсов (snapshot ``shipment_items.kp_id``)."""
    cur.execute(
        """
        SELECT COALESCE(SUM(si.qty), 0)
        FROM shipment_items si
        JOIN shipments s ON s.id = si.shipment_id
        WHERE si.kp_id = ?
          AND si.item_type = ?
          AND s.status = ?
        """,
        (int(kp_id), ShipmentItemType.PLATE.value, ShipmentStatus.DONE.value),
    )
    return int(cur.fetchone()[0] or 0)
