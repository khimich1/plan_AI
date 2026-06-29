#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read-only plate queries (A1 slice)."""

from __future__ import annotations

import sqlite3
from typing import Dict, List

from core.kp_db_common import DEFAULT_DB, _connect

def get_remaining_plates_for_kp(kp_id: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список оставшихся (невыполненных) плит для КП.
    
    Простыми словами:
    - Возвращает все плиты, которые ещё не выполнены для данного КП
    
    Аргументы:
        kp_id: номер КП
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT * FROM kp_plates 
            WHERE kp_id = ? AND qty > 0
            ORDER BY position_number
        ''', (kp_id,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


# =============================================================================
# ФУНКЦИИ ДЛЯ УПРАВЛЕНИЯ СТАТУСАМИ ПЛИТ
# =============================================================================

def get_completed_plates_for_kp(kp_id: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список выполненных плит для КП.
    
    Простыми словами:
    - Возвращает все плиты, которые уже выполнены для данного КП
    
    Аргументы:
        kp_id: номер КП
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о выполненных плитах
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT * FROM completed_plates 
            WHERE kp_id = ?
            ORDER BY completed_date, production_day
        ''', (kp_id,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_completed_plates_stats(db_path: str = DEFAULT_DB) -> Dict:
    """
    Получает статистику по выполненным плитам.
    
    Простыми словами:
    - Считает общее количество выполненных плит
    - Считает сколько КП затронуто
    - Возвращает сводку
    
    Возвращает:
        Словарь со статистикой
    """
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        
        # Общее количество записей
        cur.execute('SELECT COUNT(*) FROM completed_plates')
        total_records = cur.fetchone()[0]
        
        # Общее количество плит (сумма qty)
        cur.execute('SELECT SUM(qty) FROM completed_plates')
        result = cur.fetchone()
        total_qty = result[0] if result[0] else 0
        
        # Количество уникальных КП
        cur.execute('SELECT COUNT(DISTINCT kp_id) FROM completed_plates')
        kp_count = cur.fetchone()[0]
        
        # Количество дней производства
        cur.execute('SELECT COUNT(DISTINCT production_day) FROM completed_plates')
        days_count = cur.fetchone()[0]
        
        return {
            'total_records': total_records,
            'total_plates': total_qty,
            'kp_count': kp_count,
            'days_count': days_count
        }
        
    finally:
        conn.close()


def get_completed_plates_by_day(production_day: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все выполненные плиты за конкретный день производства.
    
    Аргументы:
        production_day: номер дня производства
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT cp.*, ko.customer_name
            FROM completed_plates cp
            LEFT JOIN KP_offers ko ON cp.kp_id = ko.kp_id
            WHERE cp.production_day = ?
            ORDER BY cp.kp_id, cp.plate_name
        ''', (production_day,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_all_plates_in_production(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все плиты в производстве (из таблицы kp_plates).
    
    Простыми словами:
    - Возвращает все плиты, которые ещё не выполнены
    - Объединяет с информацией о КП (клиент, дата)
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                p.id,
                p.kp_id,
                p.position_number,
                p.plate_name,
                p.length_m,
                p.width_m,
                p.load_class,
                p.qty,
                p.unit_weight,
                p.total_weight,
                p.discounted_price,
                p.status as plate_status,
                p.plan_id,
                ko.customer_name,
                ko.execution_terms,
                p.concrete_grade AS concrete_grade
            FROM kp_plates p
            LEFT JOIN KP_offers ko ON p.kp_id = ko.kp_id
            LEFT JOIN kp_meta m ON p.kp_id = m.kp_id
            WHERE p.qty > 0 AND p.status = 'в производстве' AND (m.status IS NULL OR m.status = 'в работе')
            ORDER BY p.kp_id, p.position_number
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_all_completed_plates(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все выполненные плиты (из таблицы completed_plates).
    
    Простыми словами:
    - Возвращает все плиты, которые уже выполнены
    - Объединяет с информацией о КП (клиент)
    
    Возвращает:
        Список словарей с информацией о выполненных плитах
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                cp.id,
                cp.kp_id,
                cp.plate_name,
                cp.length_m,
                cp.width_m,
                cp.load_class,
                cp.qty,
                cp.completed_date,
                cp.production_day,
                ko.customer_name,
                ko.execution_terms
            FROM completed_plates cp
            LEFT JOIN KP_offers ko ON cp.kp_id = ko.kp_id
            ORDER BY cp.completed_date DESC, cp.kp_id, cp.plate_name
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


