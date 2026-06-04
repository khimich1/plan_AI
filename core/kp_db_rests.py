#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate rests persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Dict, List, Optional

from core.kp_db_common import DEFAULT_DB, _connect



# ==================== ФУНКЦИИ ДЛЯ ОСТАТКОВ ПЛИТ ====================

def create_plate_rest(
    kp_id: int,
    source_plate_name: str,
    rest_width_mm: int,
    length_m: float,
    production_day: int,
    qty: int = 1,
    db_path: str = DEFAULT_DB,
    *,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> int:
    """
    Создает запись об остатке плиты.

    Простыми словами:
    - При продольном резе плиты образуется остаток
    - Эта функция сохраняет информацию об остатке в БД
    - Остаток можно использовать для других заказов

    Аргументы:
        kp_id: номер КП, при выполнении которого образовался остаток
        source_plate_name: имя исходной плиты (из которой вырезали)
        rest_width_mm: ширина остатка в мм
        length_m: длина остатка в метрах
        production_day: номер дня производства
        qty: количество остатков (по умолчанию 1)
        db_path: путь к базе данных
        _external_conn: если задано — функция работает в существующей
            транзакции переданного соединения (P0/P6). Без commit/rollback/close.

    Возвращает:
        ID созданной записи или 0 при ошибке
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
        created_date = datetime.now().strftime('%d.%m.%Y')

        cur.execute('''
            INSERT INTO plate_rests (
                kp_id, source_plate_name, rest_width_mm, length_m,
                qty, status, created_date, production_day
            ) VALUES (?, ?, ?, ?, ?, 'available', ?, ?)
        ''', (
            kp_id,
            source_plate_name,
            rest_width_mm,
            length_m,
            qty,
            created_date,
            production_day,
        ))

        rest_id = cur.lastrowid
        if own_conn:
            conn.commit()
        print(f"[DB] ✅ Создан остаток #{rest_id}: {rest_width_mm}мм x {length_m}м (КП #{kp_id})")
        return rest_id

    except Exception as e:
        if own_conn:
            print(f"[DB] ❌ Ошибка при создании остатка: {e}")
            conn.rollback()
            return 0
        # Внешняя транзакция: пробрасываем для отката caller'ом
        raise

    finally:
        if own_conn:
            conn.close()


def get_available_rests(
    kp_id: int = None,
    db_path: str = DEFAULT_DB
) -> List[Dict]:
    """
    Возвращает список доступных остатков.
    
    Простыми словами:
    - Получает все остатки со статусом 'available'
    - Можно фильтровать по номеру КП
    - Возвращает список словарей с информацией об остатках
    
    Аргументы:
        kp_id: номер КП для фильтрации (None = все КП)
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией об остатках
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        if kp_id:
            cur.execute('''
                SELECT 
                    pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                    pr.length_m, pr.qty, pr.status, pr.created_date,
                    pr.production_day, ko.customer_name
                FROM plate_rests pr
                LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
                WHERE pr.status = 'available' AND pr.kp_id = ?
                ORDER BY pr.created_date DESC, pr.rest_width_mm DESC
            ''', (kp_id,))
        else:
            cur.execute('''
                SELECT 
                    pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                    pr.length_m, pr.qty, pr.status, pr.created_date,
                    pr.production_day, ko.customer_name
                FROM plate_rests pr
                LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
                WHERE pr.status = 'available'
                ORDER BY pr.created_date DESC, pr.rest_width_mm DESC
            ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def mark_rest_as_used(
    rest_id: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Помечает остаток как использованный.
    
    Простыми словами:
    - Когда остаток используется во вторичном резе
    - Его статус меняется на 'used'
    - Он больше не будет показываться как доступный
    
    Аргументы:
        rest_id: ID остатка в БД
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'used', used_date = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} помечен как использованный")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже использован")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при обновлении остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def complete_plate_rest(
    rest_id: int,
    production_day: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Помечает остаток как выполненный (произведённый).
    
    Простыми словами:
    - Когда остаток изготавливается как отдельная плита
    - Его статус меняется на 'completed'
    - Записывается день производства
    
    Аргументы:
        rest_id: ID остатка в БД
        production_day: номер дня производства
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'completed', used_date = ?, production_day = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, production_day, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} выполнен (день {production_day})")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже обработан")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при завершении остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def discard_plate_rest(
    rest_id: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Списывает остаток (брак или утилизация).
    
    Простыми словами:
    - Когда остаток больше не нужен (брак, повреждение)
    - Его статус меняется на 'discarded'
    - Остаток удаляется из доступных
    
    Аргументы:
        rest_id: ID остатка в БД
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'discarded', used_date = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} списан")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже обработан")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при списании остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def get_all_plate_rests(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все остатки (для статистики и отчётов).
    
    Возвращает:
        Список словарей с информацией обо всех остатках
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                pr.length_m, pr.qty, pr.status, pr.created_date,
                pr.used_date, pr.production_day, ko.customer_name
            FROM plate_rests pr
            LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
            ORDER BY pr.created_date DESC, pr.status, pr.rest_width_mm DESC
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def _fetch_available_rests_candidates(
    cur: sqlite3.Cursor,
    length_m: float,
    width_mm: int,
) -> list[sqlite3.Row]:
    """SQL: available rests >= required dimensions, best match first."""
    cur.execute(
        """
            SELECT
                pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                pr.length_m, pr.qty, pr.status, pr.created_date,
                pr.production_day, ko.customer_name
            FROM plate_rests pr
            LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
            WHERE pr.status = 'available'
              AND pr.length_m >= ?
              AND pr.rest_width_mm >= ?
            ORDER BY
                CASE
                    WHEN pr.length_m = ? AND pr.rest_width_mm = ? THEN 0
                    WHEN pr.length_m = ? THEN 1
                    WHEN pr.rest_width_mm = ? THEN 2
                    ELSE 3
                END,
                (pr.length_m - ?) + (pr.rest_width_mm - ?) / 1000.0
        """,
        (length_m, width_mm, length_m, width_mm, length_m, width_mm, length_m, width_mm),
    )
    return list(cur.fetchall())


def find_matching_rests(
    length_m: float,
    width_mm: int,
    qty_needed: int,
    db_path: str = DEFAULT_DB,
) -> List[Dict]:
    """
    Ищет остатки для получения плиты нужного размера (backward-compatible facade).

    Оркестрация — RestMatchingService (app.services.rest_matching_service).
    """
    from core.rest_matching_service import RestMatchingService

    return RestMatchingService.find_matching_rests(
        length_m, width_mm, qty_needed, db_path
    )

