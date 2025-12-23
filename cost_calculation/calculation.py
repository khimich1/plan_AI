#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсинг названия плиты и расчет себестоимости
"""

import re
from typing import Optional, Dict
from .db import (
    init_cost_schema, init_default_constants,
    get_constant, get_concrete_norms, get_reinforcement_norms, get_izoform_norm,
    get_kef
)
import core.config_and_data as cfg


def parse_plate_name(plate_name: str) -> Optional[Dict]:
    """
    Парсит название плиты и возвращает параметры
    
    Поддерживает все форматы:
      - "ПБ 17-12-6" (базовый)
      - "ПБ 56-3,2-8п" (с запятой - дробное число)
      - "ПБ 56-32-8п" (ширина 32 дм = 3.2 м)
      - "ПБ 56-3.2-8" (с точкой)
      - "Плиты ПБ 78-12-8п" (с префиксом)
      - "ПК 80-12-8" (без буквы п)
    
    Returns:
        {
            'length_dm': 56,
            'length_m': 5.6,
            'width_dm': 32,  # или 3.2 если дробное
            'width_m': 3.2,
            'load_code': 8,
            'concrete_grade': 'М400'
        }
    """
    from core.config_and_data import parse_load_code_from_name
    
    plate_name = plate_name.strip()
    
    # Нормализуем: заменяем запятые на точки
    plate_name_normalized = plate_name.replace(',', '.')
    
    # Паттерн для парсинга с поддержкой дробных чисел: "ПБ 56-3.2-8п"
    # Также поддерживает: "Плиты ПБ", "ПК" вместо "ПБ"
    match = re.search(
        r'плит[аы]?\s*п[бк]\s*([\d\.]+)\s*-\s*([\d\.]+)',
        plate_name_normalized,
        re.IGNORECASE
    )
    
    if not match:
        # Вариант без префикса: "ПБ 56-32-8п"
        match = re.search(
            r'\bп[бк]\s*([\d\.]+)\s*-\s*([\d\.]+)',
            plate_name_normalized,
            re.IGNORECASE
        )
    
    if not match:
        return None
    
    try:
        length_dm = float(match.group(1))
        width_dm = float(match.group(2))
    except (ValueError, IndexError):
        return None
    
    # Парсим нагрузку (поддерживает "8п", "12,5п" и т.д.)
    load_code = parse_load_code_from_name(plate_name, default=8)
    
    # Переводим в метры
    length_m = length_dm / 10.0
    width_m = width_dm / 10.0
    
    # Формируем width_dm для БД запросов (может быть целым или дробным)
    width_dm_for_db = int(width_dm) if abs(width_dm - round(width_dm)) < 0.01 else width_dm
    
    # Определяем марку бетона по нагрузке
    concrete_grade = 'М500' if load_code >= 12 else 'М400'
    
    return {
        'length_dm': int(length_dm) if abs(length_dm - round(length_dm)) < 0.01 else length_dm,
        'length_m': round(length_m, 3),
        'width_dm': width_dm_for_db,
        'width_m': round(width_m, 3),
        'load_code': load_code,
        'concrete_grade': concrete_grade
    }


def get_plate_volume_from_db(length_dm: int, width_dm: int, load_code: int, db_path: str) -> Optional[float]:
    """Получает объем плиты из БД (загружен из Excel)"""
    import sqlite3
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT volume_m3 FROM plate_volumes
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def calculate_plate_volume(length_m: float, width_m: float) -> float:
    """
    Рассчитывает объем плиты в м³ (fallback, если нет в БД)
    
    Формула: V = L × W × H
    где H = 0.22 м (стандартная высота плиты ПБ)
    """
    height_m = 0.22
    return length_m * width_m * height_m


def calculate_plate_cost(plate_name: str, db_path: str = None) -> Optional[Dict]:
    """
    Рассчитывает себестоимость плиты по названию
    
    Args:
        plate_name: название плиты (например, "ПБ 17-12-6")
        db_path: путь к БД (по умолчанию из cfg.PRICE_DB_PATH)
    
    Returns:
        {
            'plate_name': 'ПБ 17-12-6',
            'parameters': {...},
            'components': {
                'concrete': 1708.91,
                'reinforcement': 213.28,
                'loops': 286.0,
                'izoform': 11.63
            },
            'total_cost': 2219.82,
            'breakdown': {...}
        }
    """
    if db_path is None:
        db_path = cfg.PRICE_DB_PATH
    
    init_cost_schema(db_path)
    
    params = parse_plate_name(plate_name)
    if not params:
        return None
    
    length_m = params['length_m']
    width_m = params['width_m']
    length_dm = params['length_dm']
    width_dm = params['width_dm']
    load_code = params['load_code']
    concrete_grade = params['concrete_grade']
    
    # Получаем объем из БД (загружен из Excel)
    volume_m3 = get_plate_volume_from_db(length_dm, width_dm, load_code, db_path)
    if volume_m3 is None:
        # Fallback: рассчитываем по формуле
        volume_m3 = calculate_plate_volume(length_m, width_m)
    
    # Получаем стоимость бетона из БД (загружена из Excel)
    concrete_cost = get_concrete_cost_from_db(length_dm, width_dm, load_code, db_path)
    if concrete_cost is None:
        # Fallback: рассчитываем по формуле
        concrete_cost = calculate_concrete_cost(volume_m3, concrete_grade, db_path)
    
    # Получаем стоимость армирования из БД (загружена из Excel)
    reinforcement_cost = get_reinforcement_cost_from_db(length_dm, width_dm, load_code, db_path)
    if reinforcement_cost is None:
        # Fallback: рассчитываем по формуле
        reinforcement_cost = calculate_reinforcement_cost(volume_m3, load_code, db_path)
    
    # Петли - берем из констант
    loops_cost = calculate_loops_cost(load_code, db_path)
    
    # Получаем стоимость изоформа из БД (загружена из Excel)
    izoform_cost = get_izoform_cost_from_db(length_dm, width_dm, load_code, db_path)
    if izoform_cost is None:
        # Fallback: рассчитываем по формуле
        izoform_cost = calculate_izoform_cost(volume_m3, db_path)
    
    # Прямые затраты (как в листе "Стоимость" Excel)
    # Формула: R4 = H4 + L4 + P4 + Q4 (Бетон + Армирование + Петли + Изоформ)
    direct_cost = concrete_cost + reinforcement_cost + loops_cost + izoform_cost
    
    # Применяем КЭФ (коэффициент накладных расходов, как в листе "Себестоимость" Excel)
    # Формула Excel: M3 = L3 × КЭФ (накладные = прямые × КЭФ)
    # Формула Excel: E3 = Прямые + M3 (полная = прямые + накладные)
    # Ищем КЭФ для конкретной плиты из БД (если был импортирован из Excel)
    kef = get_kef(length_dm=length_dm, width_dm=width_dm, load_code=load_code, db_path=db_path)
    
    # Накладные расходы = Прямые × КЭФ (как в Excel: M3 = L3 × КЭФ)
    overhead_cost = direct_cost * kef
    
    # Полная себестоимость = Прямые + Накладные (как в Excel: E3 = Прямые + M3)
    # Или: Полная = Прямые × (1 + КЭФ)
    full_cost_with_kef = direct_cost + overhead_cost
    
    return {
        'plate_name': plate_name,
        'parameters': params,
        'volume_m3': volume_m3,
        'components': {
            'concrete': round(concrete_cost, 2),
            'reinforcement': round(reinforcement_cost, 2),
            'loops': round(loops_cost, 2),
            'izoform': round(izoform_cost, 2)
        },
        'direct_cost': round(direct_cost, 2),           # Прямые затраты
        'overhead_cost': round(overhead_cost, 2),       # Накладные расходы
        'full_cost_with_kef': round(full_cost_with_kef, 2),  # Полная себестоимость с КЭФ
        'kef': kef,                                     # Значение КЭФ
        'total_cost': round(direct_cost, 2),           # Для обратной совместимости
        'breakdown': get_cost_breakdown(
            length_dm, width_dm, load_code, volume_m3, concrete_grade, db_path
        )
    }


def get_concrete_cost_from_db(length_dm: int, width_dm: int, load_code: int, db_path: str) -> Optional[float]:
    """Получает стоимость бетона из БД (загружена из Excel)"""
    import sqlite3
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT concrete_cost FROM concrete_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def calculate_concrete_cost(volume_m3: float, concrete_grade: str, db_path: str) -> float:
    """Рассчитывает стоимость бетона (fallback, если нет в БД)"""
    norms = get_concrete_norms(concrete_grade, db_path)
    if not norms:
        return 0.0
    
    # Получаем цены материалов (все в руб/кг)
    cement_price = get_constant('cement_price_per_kg', db_path) or 0
    sand_price = get_constant('sand_price_per_kg', db_path) or 0
    gravel_price = get_constant('gravel_price_per_kg', db_path) or 0
    polyplast_price = get_constant('polyplast_price_per_kg', db_path) or 0
    
    # Рассчитываем количество материалов (все в кг)
    cement_kg = norms['cement_kg_per_m3'] * volume_m3
    sand_kg = norms['sand_kg_per_m3'] * volume_m3
    gravel_kg = norms['gravel_kg_per_m3'] * volume_m3
    polyplast_kg = norms.get('polyplast_kg_per_m3', 0.0) * volume_m3
    
    # Рассчитываем стоимость каждого компонента
    cement_cost = cement_kg * cement_price
    sand_cost = sand_kg * sand_price
    gravel_cost = gravel_kg * gravel_price
    polyplast_cost = polyplast_kg * polyplast_price
    
    return cement_cost + sand_cost + gravel_cost + polyplast_cost


def get_reinforcement_cost_from_db(length_dm: int, width_dm: int, load_code: int, db_path: str) -> Optional[float]:
    """Получает стоимость армирования из БД (загружена из Excel)
    
    В БД хранятся:
    - wire_kg - количество проволоки в кг
    - cable_cost - стоимость каната в рублях
    
    Возвращает: wire_kg × цена_проволоки + cable_cost
    """
    import sqlite3
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT wire_kg, cable_cost FROM reinforcement_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        row = cur.fetchone()
        if not row:
            return None
        
        wire_kg = row[0] if row[0] else 0
        cable_cost = row[1] if row[1] else 0
        
        # Рассчитываем стоимость проволоки
        wire_price = get_constant('wire_price_per_kg', db_path) or 0
        wire_cost = wire_kg * wire_price
        
        # Итого: проволока + канат
        return wire_cost + cable_cost
    finally:
        conn.close()


def calculate_reinforcement_cost(volume_m3: float, load_code: int, db_path: str) -> float:
    """Рассчитывает стоимость армирования (fallback, если нет в БД)"""
    norms = get_reinforcement_norms(load_code, db_path)
    if not norms:
        return 0.0
    
    wire_price = get_constant('wire_price_per_kg', db_path) or 0
    
    wire_kg = norms['wire_kg_per_m3'] * volume_m3
    wire_cost = wire_kg * wire_price
    
    cable_cost = norms['cable_cost_per_m3'] * volume_m3
    
    return wire_cost + cable_cost


def calculate_loops_cost(load_code: int, db_path: str) -> float:
    """Рассчитывает стоимость петель
    
    По умолчанию используется 2 петли на плиту (из Excel: 8 стержней = 4 петли по 2 стержня)
    """
    loop_price = get_constant('loop_d18_price', db_path) or 0
    loops_quantity = get_constant('loops_per_plate', db_path) or 2.0
    return loop_price * loops_quantity


def get_izoform_cost_from_db(length_dm: int, width_dm: int, load_code: int, db_path: str) -> Optional[float]:
    """Получает стоимость изоформа из БД (загружена из Excel)"""
    import sqlite3
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT izoform_cost FROM izoform_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        row = cur.fetchone()
        return float(row[0]) if row else None
    finally:
        conn.close()


def calculate_izoform_cost(volume_m3: float, db_path: str) -> float:
    """Рассчитывает стоимость изоформа (fallback, если нет в БД)"""
    izoform_norm = get_izoform_norm(volume_m3, db_path)
    if not izoform_norm:
        return 0.0
    
    izoform_price = get_constant('izoform_price_per_kg', db_path) or 0
    return izoform_norm * izoform_price


def get_cost_breakdown(length_dm: int, width_dm: int, load_code: int, volume_m3: float, 
                       concrete_grade: str, db_path: str) -> Dict:
    """Возвращает детальную разбивку себестоимости"""
    import sqlite3
    
    breakdown = {
        'concrete': {
            'volume_m3': volume_m3,
            'grade': concrete_grade,
        },
        'reinforcement': {
            'load_code': load_code,
        }
    }
    
    # Получаем данные из БД
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Армирование
        cur.execute("""
            SELECT wire_kg, cable_cost FROM reinforcement_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        row = cur.fetchone()
        if row:
            wire_kg = row[0] if row[0] else 0
            cable_cost = row[1] if row[1] else 0
            breakdown['reinforcement']['wire_kg'] = round(wire_kg, 3)
            breakdown['reinforcement']['cable_cost'] = round(cable_cost, 2)
            
            # Рассчитываем стоимость проволоки для отображения
            wire_price = get_constant('wire_price_per_kg', db_path) or 0
            breakdown['reinforcement']['wire_cost'] = round(wire_kg * wire_price, 2)
        else:
            # Fallback
            norms_reinforcement = get_reinforcement_norms(load_code, db_path)
            if norms_reinforcement:
                wire_kg = norms_reinforcement['wire_kg_per_m3'] * volume_m3
                cable_cost = norms_reinforcement['cable_cost_per_m3'] * volume_m3
                breakdown['reinforcement']['wire_kg'] = round(wire_kg, 3)
                breakdown['reinforcement']['cable_cost'] = round(cable_cost, 2)
                
                # Рассчитываем стоимость проволоки для отображения
                wire_price = get_constant('wire_price_per_kg', db_path) or 0
                breakdown['reinforcement']['wire_cost'] = round(wire_kg * wire_price, 2)
            else:
                breakdown['reinforcement']['wire_kg'] = 0
                breakdown['reinforcement']['cable_cost'] = 0
                breakdown['reinforcement']['wire_cost'] = 0
        
        # Изоформ
        cur.execute("""
            SELECT izoform_kg FROM izoform_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        row = cur.fetchone()
        if row:
            breakdown['izoform_kg'] = round(row[0], 4) if row[0] else 0
    finally:
        conn.close()
    
    # Бетон - рассчитываем по нормам
    norms_concrete = get_concrete_norms(concrete_grade, db_path)
    if norms_concrete:
        breakdown['concrete']['cement_kg'] = round(norms_concrete['cement_kg_per_m3'] * volume_m3, 3)
        breakdown['concrete']['sand_kg'] = round(norms_concrete['sand_kg_per_m3'] * volume_m3, 3)
        breakdown['concrete']['gravel_kg'] = round(norms_concrete['gravel_kg_per_m3'] * volume_m3, 3)
        breakdown['concrete']['polyplast_kg'] = round(norms_concrete.get('polyplast_kg_per_m3', 0.0) * volume_m3, 3)
    else:
        breakdown['concrete']['cement_kg'] = 0
        breakdown['concrete']['sand_kg'] = 0
        breakdown['concrete']['gravel_kg'] = 0
        breakdown['concrete']['polyplast_kg'] = 0
    
    return breakdown

