#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate completion persistence (A1 slice) + KP status on SGP."""

from __future__ import annotations

import sqlite3
from typing import Dict, List, Optional

from app.domain.enums import KpStatus
from core.kp_db_common import DEFAULT_DB, _connect


def move_plates_to_completed(
    kp_id: int,
    plates_to_complete: List[Dict],
    production_day: int,
    db_path: str = DEFAULT_DB,
    plan_ids: Optional[List[str]] = None,
    allow_cross_kp: bool = False,
    *,
    actor: str | None = None,
    return_unmoved: bool = False,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> int | tuple[int, list]:
    """Facade over PlateCompletionService (A2)."""
    from core.plate_completion_service import PlateCompletionService

    return PlateCompletionService.move_plates_to_completed(
        kp_id,
        plates_to_complete,
        production_day,
        db_path,
        plan_ids,
        allow_cross_kp,
        actor=actor,
        return_unmoved=return_unmoved,
        _external_conn=_external_conn,
    )


def freeze_ordered_qty_if_needed(
    cur: sqlite3.Cursor,
    kp_id: int,
) -> int | None:
    """Freeze ``kp_meta.ordered_qty`` once (M for N/M badge).

    M = SUM(kp_plates.qty) + SUM(completed_plates.qty WHERE kp_id=?) at freeze time.
    Returns frozen value (existing or newly set), or None if kp_meta row missing.
    """
    cur.execute(
        "SELECT ordered_qty FROM kp_meta WHERE kp_id = ?",
        (kp_id,),
    )
    row = cur.fetchone()
    if row is None:
        return None
    if row[0] is not None:
        return int(row[0])

    cur.execute(
        "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = ?",
        (kp_id,),
    )
    in_kp = int(cur.fetchone()[0] or 0)
    cur.execute(
        "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = ?",
        (kp_id,),
    )
    on_sgp = int(cur.fetchone()[0] or 0)
    ordered = in_kp + on_sgp
    cur.execute(
        "UPDATE kp_meta SET ordered_qty = ? WHERE kp_id = ? AND ordered_qty IS NULL",
        (ordered, kp_id),
    )
    return ordered


def check_and_update_kp_completion(
    kp_id: int,
    db_path: str = DEFAULT_DB,
    *,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> bool:
    """
    Проверяет, все ли плиты КП на СГП и привязаны.

    Если ``remaining_in_kp_plates == 0`` и есть linked qty на СГП —
    ставит статус ``На СГП``. Иначе (после unlink и т.п.) — ``в работе``.

    Не ставит ``выполнено`` (отгрузка — OUT of MVP).

    Возвращает:
        True если КП полностью на СГП (linked), False иначе
    """
    own_conn = _external_conn is None
    if own_conn:
        conn = _connect(db_path)
    else:
        conn = _external_conn

    try:
        if own_conn:
            conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()

        freeze_ordered_qty_if_needed(cur, kp_id)

        cur.execute("SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = ?", (kp_id,))
        remaining = int(cur.fetchone()[0] or 0)

        cur.execute(
            """
            SELECT COALESCE(SUM(qty), 0) FROM completed_plates
            WHERE kp_id = ?
            """,
            (kp_id,),
        )
        linked_on_sgp = int(cur.fetchone()[0] or 0)

        if remaining == 0 and linked_on_sgp > 0:
            cur.execute(
                "UPDATE kp_meta SET status = ? WHERE kp_id = ?",
                (KpStatus.ON_SGP.value, kp_id),
            )
            if own_conn:
                conn.commit()
            print(f"[DB] КП #{kp_id} полностью на СГП. Статус обновлён.")
            return True

        # Есть потребность или нет linked СГП — держим «в работе»
        # (не трогаем «в архиве» / другие статусы, если плит ещё нет и СГП пуст)
        cur.execute("SELECT status FROM kp_meta WHERE kp_id = ?", (kp_id,))
        meta = cur.fetchone()
        current = meta[0] if meta else None
        if current == KpStatus.ON_SGP.value and remaining > 0:
            cur.execute(
                "UPDATE kp_meta SET status = ? WHERE kp_id = ?",
                (KpStatus.IN_WORK.value, kp_id),
            )
            if own_conn:
                conn.commit()
        elif remaining == 0 and linked_on_sgp == 0 and current == KpStatus.ON_SGP.value:
            cur.execute(
                "UPDATE kp_meta SET status = ? WHERE kp_id = ?",
                (KpStatus.IN_WORK.value, kp_id),
            )
            if own_conn:
                conn.commit()
        return False

    finally:
        if own_conn:
            conn.close()
