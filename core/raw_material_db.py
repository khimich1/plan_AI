#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль для работы с данными о стоимости сырья и производственных расходов
"""
import os
import sqlite3
from typing import Optional

# Путь к базе данных в корне проекта
DEFAULT_DB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'pb.db')


def init_schema(db_path: str = DEFAULT_DB) -> None:
    """Создает таблицу raw_material_costs если её нет"""
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS raw_material_costs (
                plate_name TEXT PRIMARY KEY,
                raw_material_and_production_cost REAL NOT NULL
            )
        ''')
        conn.commit()
    finally:
        conn.close()


def get_raw_material_cost(plate_name: str, db_path: str = DEFAULT_DB) -> Optional[float]:
    """
    Получает стоимость сырья и производственных расходов для указанной плиты.
    
    Args:
        plate_name: Название плиты в формате "ПБ 17-12-6"
        db_path: Путь к базе данных
    
    Returns:
        Стоимость сырья и производственных расходов или None если не найдено
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            'SELECT raw_material_and_production_cost FROM raw_material_costs WHERE plate_name = ?',
            (plate_name,)
        )
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def get_all_costs(db_path: str = DEFAULT_DB) -> dict:
    """
    Получает все данные о стоимости сырья и производственных расходов.
    
    Returns:
        Словарь {plate_name: cost}
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('SELECT plate_name, raw_material_and_production_cost FROM raw_material_costs')
        return {row[0]: float(row[1]) for row in cur.fetchall()}
    finally:
        conn.close()


def add_or_update_cost(plate_name: str, cost: float, db_path: str = DEFAULT_DB) -> None:
    """
    Добавляет или обновляет стоимость сырья и производственных расходов для плиты.
    
    Args:
        plate_name: Название плиты в формате "ПБ 17-12-6"
        cost: Стоимость сырья и производственных расходов
        db_path: Путь к базе данных
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            'INSERT OR REPLACE INTO raw_material_costs (plate_name, raw_material_and_production_cost) VALUES (?, ?)',
            (plate_name, cost)
        )
        conn.commit()
    finally:
        conn.close()


def get_statistics(db_path: str = DEFAULT_DB) -> dict:
    """
    Получает статистику по стоимости сырья и производственных расходов.
    
    Returns:
        Словарь со статистикой: count, min, max, avg
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('''
            SELECT 
                COUNT(*) as count,
                MIN(raw_material_and_production_cost) as min_cost,
                MAX(raw_material_and_production_cost) as max_cost,
                AVG(raw_material_and_production_cost) as avg_cost
            FROM raw_material_costs
        ''')
        row = cur.fetchone()
        return {
            'count': row[0],
            'min': float(row[1]) if row[1] else None,
            'max': float(row[2]) if row[2] else None,
            'avg': float(row[3]) if row[3] else None,
        }
    finally:
        conn.close()


if __name__ == '__main__':
    # Примеры использования
    print("Statistika po tablitse raw_material_costs:")
    stats = get_statistics()
    print(f"  Vsego zapisey: {stats['count']}")
    print(f"  Min stoimost: {stats['min']:.2f} rub.")
    print(f"  Max stoimost: {stats['max']:.2f} rub.")
    print(f"  Srednyaya stoimost: {stats['avg']:.2f} rub.")
    
    print("\nPrimery dannykh:")
    all_costs = get_all_costs()
    for i, (plate, cost) in enumerate(list(all_costs.items())[:5]):
        print(f"  {plate}: {cost:.2f} rub.")
    
    print("\nPoisk konkretnoy plity:")
    test_plate = "ПБ 17-12-6"
    cost = get_raw_material_cost(test_plate)
    if cost:
        print(f"  {test_plate}: {cost:.2f} rub.")
    else:
        print(f"  {test_plate}: ne naydena")

