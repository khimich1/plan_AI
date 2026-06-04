"""Shared SQLite helpers for kp_db decomposition (A1)."""

from __future__ import annotations

import os
import sqlite3
from typing import Optional

DEFAULT_DB = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "plita.db",
)


def _sqlite_casefold(value: str | None) -> str | None:
    if value is None:
        return None
    return value.casefold()


def _connect(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys = ON")
    conn.create_function("casefold", 1, _sqlite_casefold)
    return conn


def _audit_append(
    cur: sqlite3.Cursor,
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
) -> None:
    """Deprecated — use :func:`core.kp_db_audit.audit_append` or PlateAuditRepository."""
    from core.kp_db_audit import audit_append

    audit_append(
        cur,
        plate_id=plate_id,
        kp_id=kp_id,
        plate_name=plate_name,
        plan_id=plan_id,
        day_number=day_number,
        from_status=from_status,
        to_status=to_status,
        qty=qty,
        reason=reason,
        actor=actor,
    )
