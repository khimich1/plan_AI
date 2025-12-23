#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
API для получения заводской себестоимости плит ПБ

Основные функции:
- get_cost_by_plate_name(): По названию плиты
- get_cost_by_params(): По параметрам (длина, ширина)

ВАЖНО: Нагрузка определяется через get_load_code_for_plate() из config_and_data.py
"""

import sqlite3
from typing import Optional, Dict, List
from core.config_and_data import get_load_code_for_plate, make_plate_name, parse_load_code_from_name

from .db_schema import DEFAULT_DB_PATH


def get_cost_by_plate_name(
    plate_name: str,
    db_path: str = DEFAULT_DB_PATH
) -> Optional[Dict]:
    """
    Получает себестоимость плиты по её названию.
    
    Args:
        plate_name: Название плиты (например, "Плиты ПБ 71-12-10п")
        db_path: Путь к БД
    
    Returns:
        Словарь с данными о себестоимости или None
    """
    # Парсим нагрузку из названия
    load_code = parse_load_code_from_name(plate_name, default=8)
    
    # Проверка на 12.5п
    if '12,5' in plate_name or '12.5' in plate_name:
        load_code = 12.5
    
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Поиск по точному названию
        cur.execute("""
            SELECT 
                plate_name, length_dm, width_dm, load_code,
                direct_cost, overhead_cost, full_cost,
                kef, full_cost_with_kef,
                volume_m3, concrete_grade, quality_flag
            FROM factory_plate_costs
            WHERE plate_name = ?
        """, (plate_name,))
        
        row = cur.fetchone()
        if not row:
            return None
        
        # Получаем компоненты
        cur.execute("""
            SELECT component, value
            FROM factory_plate_cost_components
            WHERE plate_name = ?
        """, (plate_name,))
        
        components = {comp: val for comp, val in cur.fetchall()}
        
        return {
            'plate_name': row[0],
            'length_dm': row[1],
            'width_dm': row[2],
            'load_code': row[3],
            'direct_cost': row[4],
            'overhead_cost': row[5],
            'full_cost': row[6],
            'kef': row[7],
            'full_cost_with_kef': row[8],
            'volume_m3': row[9],
            'concrete_grade': row[10],
            'quality_flag': row[11],
            'components': components,
        }
        
    finally:
        conn.close()


def get_cost_by_params(
    length_m: float,
    width_m: float,
    db_path: str = DEFAULT_DB_PATH,
    default_load: int = 8
) -> Optional[Dict]:
    """
    Получает себестоимость плиты по параметрам (длина, ширина).
    
    ВАЖНО: Нагрузка определяется автоматически через get_load_code_for_plate().
    
    Args:
        length_m: Длина плиты в метрах
        width_m: Ширина плиты в метрах
        db_path: Путь к БД
        default_load: Нагрузка по умолчанию (если не найдена в заказе)
    
    Returns:
        Словарь с данными о себестоимости или None
    """
    # Определяем нагрузку через существующий парсер
    load_code = get_load_code_for_plate(length_m, width_m, default=default_load)
    
    # Переводим в дециметры
    length_dm = int(round(length_m * 10))
    width_dm = int(round(width_m * 10))
    
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Поиск по (длина, ширина, нагрузка)
        cur.execute("""
            SELECT 
                plate_name, length_dm, width_dm, load_code,
                direct_cost, overhead_cost, full_cost,
                kef, full_cost_with_kef,
                volume_m3, concrete_grade, quality_flag
            FROM factory_plate_costs
            WHERE length_dm = ? AND width_dm = ? AND load_code = ?
        """, (length_dm, width_dm, load_code))
        
        row = cur.fetchone()
        
        if not row:
            # Fallback: попробуем найти с другой нагрузкой (8п как дефолт)
            if load_code != 8:
                cur.execute("""
                    SELECT 
                        plate_name, length_dm, width_dm, load_code,
                        direct_cost, overhead_cost, full_cost,
                        kef, full_cost_with_kef,
                        volume_m3, concrete_grade, quality_flag
                    FROM factory_plate_costs
                    WHERE length_dm = ? AND width_dm = ? AND load_code = 8
                """, (length_dm, width_dm))
                
                row = cur.fetchone()
        
        if not row:
            return None
        
        plate_name = row[0]
        
        # Получаем компоненты
        cur.execute("""
            SELECT component, value
            FROM factory_plate_cost_components
            WHERE plate_name = ?
        """, (plate_name,))
        
        components = {comp: val for comp, val in cur.fetchall()}
        
        return {
            'plate_name': plate_name,
            'length_dm': row[1],
            'width_dm': row[2],
            'load_code': row[3],
            'direct_cost': row[4],
            'overhead_cost': row[5],
            'full_cost': row[6],
            'kef': row[7],
            'full_cost_with_kef': row[8],
            'volume_m3': row[9],
            'concrete_grade': row[10],
            'quality_flag': row[11],
            'components': components,
        }
        
    finally:
        conn.close()


def get_all_available_plates(db_path: str = DEFAULT_DB_PATH) -> List[Dict]:
    """
    Возвращает список всех плит с себестоимостью в БД.
    
    Returns:
        Список словарей с основными параметрами плит
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        cur.execute("""
            SELECT 
                plate_name, length_dm, width_dm, load_code,
                full_cost_with_kef, volume_m3, concrete_grade
            FROM factory_plate_costs
            ORDER BY length_dm, width_dm, load_code
        """)
        
        results = []
        for row in cur.fetchall():
            results.append({
                'plate_name': row[0],
                'length_dm': row[1],
                'width_dm': row[2],
                'load_code': row[3],
                'full_cost_with_kef': row[4],
                'volume_m3': row[5],
                'concrete_grade': row[6],
            })
        
        return results
        
    finally:
        conn.close()


def get_cost_breakdown(plate_name: str, db_path: str = DEFAULT_DB_PATH) -> Optional[Dict]:
    """
    Получает детальную разбивку себестоимости по компонентам.
    
    Args:
        plate_name: Название плиты
        db_path: Путь к БД
    
    Returns:
        Словарь с разбивкой или None
    """
    cost_data = get_cost_by_plate_name(plate_name, db_path)
    if not cost_data:
        return None
    
    components = cost_data['components']
    direct_cost = cost_data['direct_cost']
    
    # Рассчитываем проценты
    breakdown = {}
    for comp, value in components.items():
        pct = (value / direct_cost * 100) if direct_cost > 0 else 0
        breakdown[comp] = {
            'value': value,
            'percentage': round(pct, 1)
        }
    
    return {
        'plate_name': plate_name,
        'direct_cost': direct_cost,
        'overhead_cost': cost_data['overhead_cost'],
        'full_cost_with_kef': cost_data['full_cost_with_kef'],
        'kef': cost_data['kef'],
        'breakdown': breakdown,
    }


if __name__ == '__main__':
    # Тестирование API
    print("=== Тест API заводской себестоимости ===\n")
    
    # Тест 1: По названию
    print("[Тест 1] Поиск по названию 'Плиты ПБ 71-12-10п'")
    cost = get_cost_by_plate_name("Плиты ПБ 71-12-10п")
    if cost:
        print(f"✓ Найдено: {cost['plate_name']}")
        print(f"  Прямые затраты: {cost['direct_cost']:.2f} руб")
        print(f"  С КЭФ: {cost['full_cost_with_kef']:.2f} руб")
        print(f"  Компоненты: {list(cost['components'].keys())}")
    else:
        print("✗ Не найдено")
    
    # Тест 2: По параметрам
    print("\n[Тест 2] Поиск по параметрам (7.1м × 1.2м)")
    cost = get_cost_by_params(7.1, 1.2)
    if cost:
        print(f"✓ Найдено: {cost['plate_name']}")
        print(f"  Нагрузка: {cost['load_code']}")
        print(f"  Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
    else:
        print("✗ Не найдено")
    
    # Тест 3: Список плит
    print("\n[Тест 3] Первые 5 плит в БД")
    plates = get_all_available_plates()
    if plates:
        print(f"✓ Всего плит в БД: {len(plates)}")
        for p in plates[:5]:
            print(f"  - {p['plate_name']}: {p['full_cost_with_kef']:.2f} руб")
    else:
        print("✗ БД пуста")

