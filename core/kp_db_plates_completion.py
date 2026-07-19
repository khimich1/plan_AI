#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate completion persistence (A1 slice)."""

from __future__ import annotations

import sqlite3
from typing import Dict, List, Optional

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

def check_and_update_kp_completion(
    kp_id: int,
    db_path: str = DEFAULT_DB,
    *,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> bool:
    """
    Проверяет, все ли плиты КП выполнены.
    Если да — меняет статус КП на "выполнено".

    Простыми словами:
    - Смотрит, остались ли ещё плиты в kp_plates для данного КП
    - Если плит не осталось (все выполнены) — ставит статус "выполнено"

    Аргументы:
        kp_id: номер КП для проверки
        db_path: путь к базе данных
        _external_conn: если задано — функция работает в существующей транзакции
            переданного соединения (P0). Не делает commit/rollback и не закрывает conn.

    Возвращает:
        True если КП полностью выполнен, False если ещё есть плиты
    """
    own_conn = _external_conn is None
    if own_conn:
        conn = _connect(db_path)
    else:
        conn = _external_conn

    try:
        if own_conn:
            conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()

        # Считаем оставшиеся плиты в КП
        cur.execute('SELECT SUM(qty) FROM kp_plates WHERE kp_id = ?', (kp_id,))
        result = cur.fetchone()
        remaining = result[0] if result[0] else 0

        if remaining == 0:
            # Все плиты выполнены — обновляем статус
            cur.execute('''
                UPDATE kp_meta SET status = 'выполнено' WHERE kp_id = ?
            ''', (kp_id,))
            if own_conn:
                conn.commit()
            print(f"[DB] 🎉 КП #{kp_id} полностью выполнен! Статус обновлён.")
            return True

        return False

    finally:
        if own_conn:
            conn.close()


