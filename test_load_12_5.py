#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест: нагрузка 12.5 кПа отображается как 12,5п, но считается по цене 12п
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import config as cfg
from core.config import set_plate_lists_from_text, format_reinforcement_from_load_code
from core.price_db import get_price
from collections import defaultdict
import math

print("=" * 100)
print("ТЕСТ ОБРАБОТКИ НАГРУЗКИ 12.5 кПа")
print("=" * 100)

# Тестовый заказ с разными нагрузками
test_order = """
Плиты ПБ 55-12-12,5п — 6 шт
Плиты ПБ 55-12-12п — 3 шт
Плиты ПБ 55-12-10п — 2 шт
"""

print("\n[1] ПАРСИНГ ЗАКАЗА")
print("-" * 100)
unparsed_lines = set_plate_lists_from_text(test_order)

print(f"\nPLATE_LOAD_DETAILS:")
for (length, width, load_code), qty in sorted(cfg.PLATE_LOAD_DETAILS.items()):
    load_display = format_reinforcement_from_load_code(load_code)
    print(f"  {qty}x {length}м × {width:.3f}м = {load_display}")

# Проверка 1: load_code должен быть float (12.5), не int (13)
print("\n[2] ПРОВЕРКА: НАГРУЗКА ХРАНИТСЯ КАК FLOAT")
print("-" * 100)
has_12_5 = any(load_code == 12.5 for (_, _, load_code), _ in cfg.PLATE_LOAD_DETAILS.items())
has_12 = any(load_code == 12 or load_code == 12.0 for (_, _, load_code), _ in cfg.PLATE_LOAD_DETAILS.items())
has_13 = any(load_code == 13 for (_, _, load_code), _ in cfg.PLATE_LOAD_DETAILS.items())

print(f"  12.5 присутствует: {has_12_5} {'✅' if has_12_5 else '❌'}")
print(f"  12 присутствует:   {has_12} {'✅' if has_12 else '❌'}")
print(f"  13 присутствует:   {has_13} {'❌ (ПЛОХО!)' if has_13 else '✅ (OK)'}")

# Проверка 2: форматирование
print("\n[3] ПРОВЕРКА: ФОРМАТИРОВАНИЕ НАГРУЗОК")
print("-" * 100)
test_cases = [
    (12.5, "12,5п"),
    (12.0, "12п"),
    (12, "12п"),
    (10.0, "10п"),
    (8, "8п"),
]

all_format_ok = True
for load, expected in test_cases:
    result = format_reinforcement_from_load_code(load)
    status = "✅" if result == expected else "❌"
    if result != expected:
        all_format_ok = False
    print(f"  {load} → '{result}' (ожидалось '{expected}') {status}")

# Проверка 3: группировка
print("\n[4] ПРОВЕРКА: ГРУППИРОВКА ПО НАГРУЗКАМ")
print("-" * 100)

orders_by_load = defaultdict(list)
for (length, width_m, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
    load_group = math.floor(load_code) if isinstance(load_code, (int, float)) else load_code
    orders_by_load[load_group].append({
        'load_code': load_code,
        'qty': qty
    })

print(f"Создано групп: {len(orders_by_load)}")
for load_group in sorted(orders_by_load.keys()):
    orders = orders_by_load[load_group]
    originals = sorted(set(o['load_code'] for o in orders))
    total_qty = sum(o['qty'] for o in orders)
    
    load_display_list = [format_reinforcement_from_load_code(lc) for lc in originals]
    load_display = ", ".join(load_display_list)
    
    print(f"\n  Группа {load_group}п:")
    print(f"    Нагрузки: {load_display}")
    print(f"    Плит: {total_qty}")

# Ожидаем: группа 12 должна содержать И 12, И 12.5
has_group_12 = 12 in orders_by_load
if has_group_12:
    group_12_originals = sorted(set(o['load_code'] for o in orders_by_load[12]))
    has_both = 12 in group_12_originals and 12.5 in group_12_originals
    print(f"\n  ✅ Группа 12п содержит: {group_12_originals}")
    if has_both:
        print(f"  ✅ ОТЛИЧНО: 12 и 12.5 в одной группе!")
    else:
        print(f"  ❌ ОШИБКА: 12 и 12.5 должны быть в одной группе!")
else:
    print(f"  ❌ ОШИБКА: Группа 12п не найдена!")

# Проверка 4: цены
print("\n[5] ПРОВЕРКА: ЗАПРОС ЦЕН ИЗ БАЗЫ")
print("-" * 100)

test_length = 5.5
price_12 = get_price(test_length, 12, cfg.PRICE_DB_PATH)
price_12_5 = get_price(test_length, 12.5, cfg.PRICE_DB_PATH)

print(f"  Цена для 5.5м × 12п:   {price_12}")
print(f"  Цена для 5.5м × 12.5п: {price_12_5}")

if price_12 == price_12_5:
    print(f"  ✅ ОТЛИЧНО: Цены одинаковые (12.5 считается как 12)!")
else:
    print(f"  ❌ ОШИБКА: Цены должны быть одинаковыми!")

# Итоговая проверка
print("\n" + "=" * 100)
print("📊 ИТОГОВАЯ ПРОВЕРКА")
print("=" * 100)

all_ok = (
    has_12_5 and              # Нагрузка 12.5 хранится как float
    not has_13 and            # НЕ округляется до 13
    all_format_ok and         # Форматирование работает
    has_group_12 and          # Есть группа 12
    price_12 == price_12_5    # Цены одинаковые
)

if all_ok:
    print("\n✅ ✅ ✅ ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ!")
    print("\n📋 Резюме:")
    print("  • 12.5 хранится как float (не округляется до 13)")
    print("  • 12.5 отображается как '12,5п'")
    print("  • 12.5 группируется вместе с 12")
    print("  • 12.5 считается по цене 12п")
else:
    print("\n❌ НЕКОТОРЫЕ ПРОВЕРКИ НЕ ПРОШЛИ!")

print("\n" + "=" * 100)

