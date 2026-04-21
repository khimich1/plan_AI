#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для проверки группировки плит по нагрузке
"""
import sys
import os

# Добавляем корневую директорию в путь
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import config as cfg
from core.config import set_plate_lists_from_text
from core.optimization import optimize_with_cascading_longitudinal_cuts
from core.visualization import visualize_plan

# Тестовый заказ с разными нагрузками
test_order = """
Плиты ПБ 28-5,3-8п — 2 шт
Плиты ПБ 73-12-8п — 93 шт
Плиты ПБ 73-10,2-8п — 6 шт
Плиты ПБ 74-12-8п — 13 шт
Плиты ПБ 73-3,2-10п — 1 шт
Плиты ПБ 86-12-8п — 3 шт
Плиты ПБ 86-7,2-8п — 1 шт
Плиты ПБ 86-5,3-8п — 1 шт
Плиты ПБ 86-3,2-8п — 1 шт
Плиты ПБ 87-12-8п — 55 шт
Плиты ПБ 87-6,65-8п — 2 шт
Плиты ПБ 87-5,3-8п — 1 шт
Плиты ПБ 87-10,7-8п — 1 шт
Плиты ПБ 84-12-8п — 1 шт
Плиты ПБ 84-7,2-8п — 2 шт
Плиты ПБ 73-6,65-8п — 1 шт
Плиты ПБ 73-5,3-8п — 1 шт
Плиты ПБ 73-12-10п — 4 шт
Плиты ПБ 28-6,65-8п — 1 шт
Плиты ПБ 25-6,65-8п — 1 шт
Плиты ПБ 87-9,2-8п — 1 шт
Плиты ПБ 87-2,6-8п — 1 шт
Плиты ПБ 27-7,2-8п — 3 шт
Плиты ПБ 28-7,2-8п — 1 шт
Плиты ПБ 26-7,2-8п — 2 шт
Плиты ПБ 55-12-12,5п — 6 шт
Плиты ПБ 74-12-10п — 7 шт
Плиты ПБ 55-12-10п — 9 шт
Плиты ПБ 61-12-10п — 21 шт
Плиты ПБ 73-7,2-8п — 2 шт
"""

print("=" * 80)
print("ТЕСТ ГРУППИРОВКИ ПЛИТ ПО НАГРУЗКЕ")
print("=" * 80)

# Шаг 1: Парсим заказ
print("\n[ШАГ 1] Парсинг заказа...")
unparsed_lines = set_plate_lists_from_text(test_order)

if unparsed_lines:
    print(f"⚠️ Нераспознанные строки: {len(unparsed_lines)}")
    for line in unparsed_lines:
        print(f"  - {line}")
else:
    print("✅ Все строки успешно распознаны!")

# Шаг 2: Проверяем PLATE_LOAD_DETAILS
print(f"\n[ШАГ 2] Проверка PLATE_LOAD_DETAILS...")
print(f"Записей в PLATE_LOAD_DETAILS: {len(cfg.PLATE_LOAD_DETAILS)}")

# Группируем по нагрузкам для анализа
from collections import defaultdict
loads_summary = defaultdict(int)
for (length, width, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
    loads_summary[load_code] += qty

print(f"\n📊 Распределение по нагрузкам:")
for load_code in sorted(loads_summary.keys()):
    print(f"  • {load_code}п: {loads_summary[load_code]} плит")

total_plates = sum(loads_summary.values())
print(f"\n📦 Всего плит: {total_plates}")

# Шаг 3: Группируем по нагрузкам (как в боте)
print(f"\n[ШАГ 3] Группировка для оптимизации...")
orders_by_load = defaultdict(list)

for (length, width_m, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
    width_mm = int(round(width_m * 1000))
    orders_by_load[load_code].append({
        'length': length,
        'width': width_mm,
        'qty': qty,
        'load_code': load_code
    })

print(f"Создано {len(orders_by_load)} групп по нагрузкам:")
for load_code in sorted(orders_by_load.keys()):
    orders = orders_by_load[load_code]
    total_qty = sum(o['qty'] for o in orders)
    print(f"  • {load_code}п: {len(orders)} типов плит, {total_qty} шт")

# Шаг 4: Запускаем оптимизацию для каждой группы
print(f"\n[ШАГ 4] Оптимизация по группам...")
import core.optimization as optimization

optimization_results_by_load = {}
total_plates_all = 0
total_cost_all = 0

for load_code in sorted(orders_by_load.keys()):
    orders_2d = orders_by_load[load_code]
    print(f"\n  === Оптимизация для нагрузки {load_code}п ===")
    print(f"  Плит: {sum(o['qty'] for o in orders_2d)} шт, типов: {len(orders_2d)}")
    
    try:
        optimization_result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
        
        if optimization_result and optimization_result.get('total_plates', 0) > 0:
            optimization_results_by_load[load_code] = optimization_result
            total_plates_all += optimization_result.get('total_plates', 0)
            total_cost_all += optimization_result.get('total_cost', 0)
            
            print(f"  ✅ Результат: {optimization_result['total_plates']} плит 1.2м, "
                  f"{optimization_result.get('total_cost', 0):,.0f} ₽".replace(',', ' '))
        else:
            print(f"  ⚠️ Оптимизация не дала результата")
    except Exception as e:
        print(f"  ❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()

# Сохраняем результаты в глобальную переменную
if optimization_results_by_load:
    optimization.OPT_CASCADING_PLAN_BY_LOAD = optimization_results_by_load
    print(f"\n✅ Сохранено {len(optimization_results_by_load)} результатов оптимизации")
    
    print("\n" + "=" * 80)
    print("📊 ИТОГОВАЯ СВОДКА")
    print("=" * 80)
    for load_code in sorted(optimization_results_by_load.keys()):
        result = optimization_results_by_load[load_code]
        print(f"• {load_code}п: {result['total_plates']} плит 1.2м, "
              f"{result.get('total_cost', 0):,.0f} ₽".replace(',', ' '))
    
    print(f"\n{'='*80}")
    print(f"💰 ИТОГО: {total_plates_all} плит, {total_cost_all:,.0f} ₽".replace(',', ' '))
    print(f"{'='*80}")

# Шаг 5: Создаём визуализацию
print(f"\n[ШАГ 5] Создание визуализации...")
try:
    output_dir = 'Визуализация_Раскладки'
    result_paths = visualize_plan(output_dir)
    print(f"✅ Визуализация создана!")
    if result_paths:
        print(f"📁 Файлы сохранены в: {output_dir}/")
        for path in result_paths:
            print(f"  • {path}")
except Exception as e:
    print(f"❌ Ошибка визуализации: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
print("✅ ТЕСТ ЗАВЕРШЁН")
print("=" * 80)

