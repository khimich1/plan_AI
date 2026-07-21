#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Схема БД для хранения заводской себестоимости плит ПБ

Таблицы:
- factory_plate_costs: Полная себестоимость плиты
- factory_plate_cost_components: Детализация по компонентам
"""

import os
import sqlite3
from typing import Optional

# Путь к БД (корень проекта)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_DB_PATH = os.path.join(BASE_DIR, 'pb.db')


def init_factory_cost_schema(db_path: str = DEFAULT_DB_PATH) -> None:
    """
    Инициализирует схему БД для хранения заводской себестоимости.
    
    Идемпотентная функция - можно вызывать многократно.
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Таблица: Полная себестоимость плиты
        cur.execute("""
            CREATE TABLE IF NOT EXISTS factory_plate_costs (
                plate_name TEXT NOT NULL,
                length_dm INTEGER NOT NULL,
                width_dm INTEGER NOT NULL,
                load_code REAL NOT NULL,
                direct_cost REAL NOT NULL,
                overhead_cost REAL NOT NULL,
                full_cost REAL NOT NULL,
                kef REAL,
                full_cost_with_kef REAL,
                volume_m3 REAL,
                concrete_grade TEXT,
                quality_flag TEXT,
                source_file TEXT,
                source_sheet TEXT,
                source_row INTEGER,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (length_dm, width_dm, load_code)
            )
        """)
        
        # Таблица: Детализация себестоимости по компонентам
        cur.execute("""
            CREATE TABLE IF NOT EXISTS factory_plate_cost_components (
                plate_name TEXT NOT NULL,
                component TEXT NOT NULL CHECK(
                    component IN ('reinforcement', 'concrete', 'loops', 'izoform')
                ),
                value REAL NOT NULL,
                PRIMARY KEY (plate_name, component)
            )
        """)
        
        # Индексы для быстрого поиска
        cur.execute("""
            CREATE INDEX IF NOT EXISTS idx_factory_costs_plate_name 
            ON factory_plate_costs(plate_name)
        """)
        
        cur.execute("""
            CREATE INDEX IF NOT EXISTS idx_factory_costs_dimensions 
            ON factory_plate_costs(length_dm, width_dm)
        """)
        
        cur.execute("""
            CREATE INDEX IF NOT EXISTS idx_factory_components_plate 
            ON factory_plate_cost_components(plate_name)
        """)
        
        conn.commit()
        print("[DB SCHEMA] ✓ Схема factory_cost инициализирована")
        
    finally:
        conn.close()


def clear_factory_costs(db_path: str = DEFAULT_DB_PATH) -> int:
    """
    Очищает все данные о себестоимости (для переимпорта).
    
    Returns:
        Количество удалённых записей
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Считаем перед удалением
        cur.execute("SELECT COUNT(*) FROM factory_plate_costs")
        count = cur.fetchone()[0]
        
        # Удаляем
        cur.execute("DELETE FROM factory_plate_costs")
        cur.execute("DELETE FROM factory_plate_cost_components")
        
        conn.commit()
        print(f"[DB SCHEMA] ✓ Удалено {count} записей себестоимости")
        return count
        
    finally:
        conn.close()


def get_factory_cost_stats(db_path: str = DEFAULT_DB_PATH) -> dict:
    """
    Получает статистику по себестоимости в БД.
    
    Returns:
        Словарь со статистикой
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Общее количество плит
        cur.execute("SELECT COUNT(*) FROM factory_plate_costs")
        total_plates = cur.fetchone()[0]
        
        # Количество плит с проблемами
        cur.execute("""
            SELECT COUNT(*) FROM factory_plate_costs 
            WHERE quality_flag IS NOT NULL AND quality_flag != ''
        """)
        problem_plates = cur.fetchone()[0]
        
        # Диапазон длин
        cur.execute("""
            SELECT MIN(length_dm), MAX(length_dm) 
            FROM factory_plate_costs
        """)
        length_range = cur.fetchone()
        
        # Диапазон ширин
        cur.execute("""
            SELECT MIN(width_dm), MAX(width_dm) 
            FROM factory_plate_costs
        """)
        width_range = cur.fetchone()
        
        # Уникальные нагрузки
        cur.execute("SELECT DISTINCT load_code FROM factory_plate_costs ORDER BY load_code")
        load_codes = [row[0] for row in cur.fetchall()]
        
        return {
            'total_plates': total_plates,
            'problem_plates': problem_plates,
            'length_range_dm': length_range,
            'width_range_dm': width_range,
            'load_codes': load_codes,
        }
        
    finally:
        conn.close()


if __name__ == '__main__':
    # Инициализация при запуске модуля напрямую
    init_factory_cost_schema()
    stats = get_factory_cost_stats()
    print("\n=== Статистика factory_costs ===")
    print(f"Всего плит: {stats['total_plates']}")
    print(f"С проблемами: {stats['problem_plates']}")
    print(f"Длины (дм): {stats['length_range_dm']}")
    print(f"Ширины (дм): {stats['width_range_dm']}")
    print(f"Нагрузки: {stats['load_codes']}")

