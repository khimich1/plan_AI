"""Plate status audit log persistence (A5 — single implementation)."""

from __future__ import annotations

from sqlite3 import Cursor
from typing import Optional


def audit_append(
    cur: Cursor,
    *,
    plate_id: Optional[int],
    kp_id: int,
    plate_name: Optional[str],
    plan_id: Optional[str],
    day_number: Optional[int],
    from_status: Optional[str],
    to_status: str,
    qty: int,
    reason: str,
    actor: Optional[str],
    shipment_id: Optional[int] = None,
) -> None:
    """Insert one row into ``plate_status_log`` (same transaction as caller).

    ``shipment_id`` заполняется только для отгрузок (reason ``sgp_ship``);
    остальные вызовы оставляют NULL.
    """
    cur.execute(
        """
        INSERT INTO plate_status_log (
            plate_id, kp_id, plate_name, plan_id, day_number,
            from_status, to_status, qty, reason, actor, shipment_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            plate_id,
            int(kp_id),
            plate_name,
            plan_id,
            day_number,
            from_status,
            to_status,
            int(qty),
            reason,
            actor,
            shipment_id,
        ),
    )
