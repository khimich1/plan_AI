#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
База данных для хранения констант расчета себестоимости плит
"""

import os
import sqlite3
from typing import Optional, Dict

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_DB = os.path.join(BASE_DIR, 'pb.db')


def init_cost_schema(db_path: str = DEFAULT_DB) -> None:
    """Инициализирует схему БД для констант себестоимости"""
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Таблица констант материалов
        cur.execute("""
            CREATE TABLE IF NOT EXISTS cost_constants (
                key TEXT PRIMARY KEY,
                value REAL NOT NULL,
                unit TEXT,
                description TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Таблица норм расхода материалов на м³ бетона
        cur.execute("""
            CREATE TABLE IF NOT EXISTS concrete_norms (
                concrete_grade TEXT PRIMARY KEY,
                cement_kg_per_m3 REAL,
                sand_kg_per_m3 REAL,
                gravel_kg_per_m3 REAL,
                polyplast_kg_per_m3 REAL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Миграция: добавляем новые колонки если их нет
        try:
            cur.execute("ALTER TABLE concrete_norms ADD COLUMN polyplast_kg_per_m3 REAL DEFAULT 0.0")
        except sqlite3.OperationalError:
            pass  # Колонка уже существует
        
        # Проверяем и переименовываем старые колонки (если нужно)
        cur.execute("PRAGMA table_info(concrete_norms)")
        columns = [row[1] for row in cur.fetchall()]
        if 'sand_m3_per_m3' in columns:
            # Есть старые колонки - нужна полная миграция
            cur.execute("DROP TABLE IF EXISTS concrete_norms")
            cur.execute("""
                CREATE TABLE concrete_norms (
                    concrete_grade TEXT PRIMARY KEY,
                    cement_kg_per_m3 REAL,
                    sand_kg_per_m3 REAL,
                    gravel_kg_per_m3 REAL,
                    polyplast_kg_per_m3 REAL,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
        
        # Таблица норм расхода армирования по нагрузке
        cur.execute("""
            CREATE TABLE IF NOT EXISTS reinforcement_norms (
                load_code INTEGER PRIMARY KEY,
                wire_kg_per_m3 REAL,
                cable_cost_per_m3 REAL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Таблица норм расхода изоформа
        cur.execute("""
            CREATE TABLE IF NOT EXISTS izoform_norms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                volume_m3_min REAL,
                volume_m3_max REAL,
                izoform_kg_per_plate REAL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        conn.commit()
    finally:
        conn.close()


def init_default_constants(db_path: str = DEFAULT_DB) -> None:
    """Инициализирует БД значениями по умолчанию из анализа Excel"""
    init_cost_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Константы цен материалов (актуальные значения от 14.03.2025)
        constants = [
            ('cement_price_per_kg', 7.92, 'руб/кг', 'Цена цемента ПЦ 500 Д0'),
            ('sand_price_per_kg', 0.48, 'руб/кг', 'Цена песка'),
            ('gravel_price_per_kg', 2.25, 'руб/кг', 'Цена щебня рядового фр. 5(3)-20'),
            ('polyplast_price_per_kg', 62.93, 'руб/кг', 'Цена полипласта'),
            ('concrete_v30_price_per_m3', 0.0, 'руб/м³', 'Цена бетона В30 (рассчитывается)'),
            ('concrete_v40_price_per_m3', 0.0, 'руб/м³', 'Цена бетона В40 (рассчитывается)'),
            ('wire_price_per_kg', 80.0, 'руб/кг', 'Цена проволоки Вр 1400'),
            ('cable_price_per_kg', 80.0, 'руб/кг', 'Цена каната д 12 К7'),
            ('rebar_d10_price_per_kg', 41.52, 'руб/кг', 'Цена арматуры 500 Ф 10'),
            ('rebar_d12_price_per_kg', 39.85, 'руб/кг', 'Цена арматуры 500 Ф 12'),
            ('rebar_d14_price_per_kg', 39.44, 'руб/кг', 'Цена арматуры 500 Ф 14'),
            ('loop_d12_price', 0.0, 'руб/шт', 'Цена петли д 12'),
            ('loop_d14_price', 0.0, 'руб/шт', 'Цена петли д 14'),
            ('loop_d18_price', 286.0, 'руб/шт', 'Цена петли д 18'),
            ('loops_per_plate', 2.0, 'шт', 'Количество петель на плиту (стандартно)'),
            ('izoform_price_per_kg', 95.0, 'руб/кг', 'Цена Изоформ-Б "Экстра"'),
        ]
        
        for key, value, unit, desc in constants:
            cur.execute("""
                INSERT OR REPLACE INTO cost_constants (key, value, unit, description)
                VALUES (?, ?, ?, ?)
            """, (key, value, unit, desc))
        
        # Нормы расхода материалов на м³ бетона (кг/м³, из Excel от 14.03.2025)
        # В30 (Дорожка): Цемент 400, Песок 920, Щебень 940, Полипласт 4.00
        norms_v30 = ('В30', 400.0, 920.0, 940.0, 4.00)
        # В40 (Дорожка): Цемент 490, Песок 830, Щебень 940, Полипласт 4.30
        norms_v40 = ('В40', 490.0, 830.0, 940.0, 4.30)
        
        # Также создаем записи для старых обозначений М400/М500 для обратной совместимости
        # М400 ≈ В30, М500 ≈ В40
        norms_m400 = ('М400', 400.0, 920.0, 940.0, 4.00)
        norms_m500 = ('М500', 490.0, 830.0, 940.0, 4.30)
        
        for grade, cement, sand, gravel, polyplast in [norms_v30, norms_v40, norms_m400, norms_m500]:
            cur.execute("""
                INSERT OR REPLACE INTO concrete_norms 
                (concrete_grade, cement_kg_per_m3, sand_kg_per_m3, gravel_kg_per_m3, polyplast_kg_per_m3)
                VALUES (?, ?, ?, ?, ?)
            """, (grade, cement, sand, gravel, polyplast))
        
        # Нормы расхода армирования (примерные, нужно уточнить из Excel)
        reinforcement_norms = [
            (6, 6.0, 0.0),
            (8, 6.6, 0.0),
            (10, 7.0, 0.0),
            (12, 8.0, 0.0),
        ]
        
        for load_code, wire, cable in reinforcement_norms:
            cur.execute("""
                INSERT OR REPLACE INTO reinforcement_norms 
                (load_code, wire_kg_per_m3, cable_cost_per_m3)
                VALUES (?, ?, ?)
            """, (load_code, wire, cable))
        
        # Нормы расхода изоформа (примерные)
        izoform_norms = [
            (0.0, 0.3, 0.072),
            (0.3, 0.4, 0.1224),
            (0.4, 0.5, 0.15),
        ]
        
        for vol_min, vol_max, izoform_kg in izoform_norms:
            cur.execute("""
                INSERT INTO izoform_norms (volume_m3_min, volume_m3_max, izoform_kg_per_plate)
                VALUES (?, ?, ?)
            """, (vol_min, vol_max, izoform_kg))
        
        conn.commit()
    finally:
        conn.close()


def get_constant(key: str, db_path: str = DEFAULT_DB) -> Optional[float]:
    """Получает значение константы из БД"""
    init_cost_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT value FROM cost_constants WHERE key=?", (key,))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def get_concrete_norms(grade: str, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """Получает нормы расхода материалов для марки бетона"""
    init_cost_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT cement_kg_per_m3, sand_kg_per_m3, gravel_kg_per_m3, polyplast_kg_per_m3
            FROM concrete_norms WHERE concrete_grade=?
        """, (grade,))
        row = cur.fetchone()
        if row:
            return {
                'cement_kg_per_m3': row[0],
                'sand_kg_per_m3': row[1],
                'gravel_kg_per_m3': row[2],
                'polyplast_kg_per_m3': row[3] if len(row) > 3 else 0.0
            }
        return None
    finally:
        conn.close()


def get_reinforcement_norms(load_code: int, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """Получает нормы расхода армирования для нагрузки"""
    init_cost_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT wire_kg_per_m3, cable_cost_per_m3
            FROM reinforcement_norms WHERE load_code=?
        """, (load_code,))
        row = cur.fetchone()
        if row:
            return {
                'wire_kg_per_m3': row[0],
                'cable_cost_per_m3': row[1]
            }
        return None
    finally:
        conn.close()


def get_izoform_norm(volume_m3: float, db_path: str = DEFAULT_DB) -> Optional[float]:
    """Получает норму расхода изоформа для объема плиты"""
    init_cost_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT izoform_kg_per_plate FROM izoform_norms
            WHERE ? >= volume_m3_min AND ? < volume_m3_max
            ORDER BY volume_m3_min LIMIT 1
        """, (volume_m3, volume_m3))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()

