#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для сравнения OLD и NEW моделей оптимизации.

Запускает 3-5 реальных наборов заказов с двумя конфигурациями:
- OLD: unused_penalty=0.5, reuse_bonus=-500 (старое поведение)
- NEW: unused_penalty=0.15, reuse_bonus=0 (новое поведение)

Для каждого набора заказов выводит:
- Количество первичных и вторичных резов
- Суммарные отходы
- Неиспользованные остатки
- Итоговая стоимость по модели
"""

import sys
import os
from pathlib import Path

# Добавляем корень проекта в sys.path
TESTS_DIR = Path(__file__).parent
PROJECT_ROOT = TESTS_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

# ВАЖНО: Настраиваем кодировку для Windows консоли (чтобы кириллица работала)
if sys.platform == 'win32':
    # Пытаемся установить UTF-8 для консоли
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        # Если не получилось, пытаемся через другой способ
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

from core.optimization import (
    optimize_with_cascading_longitudinal_cuts,
    OptimizationConfig,
    OLD_CONFIG,
    DEFAULT_CONFIG as NEW_CONFIG
)


# ==================== ТЕСТОВЫЕ НАБОРЫ ЗАКАЗОВ ====================

# Набор 1: Типичный заказ (10×ПБ 62-12-8п + 1×ПБ 36-3-8п)
TEST_CASE_1 = {
    'name': 'Набор 1: Типичный заказ (10×ПБ 62-12-8п + 1×ПБ 36-3-8п)',
    'orders_2d': [
        {'length': 6.2, 'width': 1200, 'qty': 10},  # ПБ 62-12-8п
        {'length': 3.6, 'width': 320, 'qty': 1},    # ПБ 36-3-8п
    ]
}

# Набор 2: Много узких плит (15×ПБ 56-3-8п + 5×ПБ 78-7-8п)
TEST_CASE_2 = {
    'name': 'Набор 2: Много узких плит (15×ПБ 56-3-8п + 5×ПБ 78-7-8п)',
    'orders_2d': [
        {'length': 5.6, 'width': 320, 'qty': 15},  # ПБ 56-3-8п
        {'length': 7.8, 'width': 720, 'qty': 5},   # ПБ 78-7-8п
    ]
}

# Набор 3: Смешанный заказ с разными длинами
TEST_CASE_3 = {
    'name': 'Набор 3: Смешанный (8×ПБ 62-3-8п + 6×ПБ 62-8-8п + 4×ПБ 78-10-8п)',
    'orders_2d': [
        {'length': 6.2, 'width': 320, 'qty': 8},   # ПБ 62-3-8п
        {'length': 6.2, 'width': 860, 'qty': 6},   # ПБ 62-8-8п
        {'length': 7.8, 'width': 1080, 'qty': 4},  # ПБ 78-10-8п
    ]
}

# Набор 4: Заказ с широкими плитами (12×ПБ 68-8-8п + 3×ПБ 68-4-8п)
TEST_CASE_4 = {
    'name': 'Набор 4: Широкие плиты (12×ПБ 68-8-8п + 3×ПБ 68-4-8п)',
    'orders_2d': [
        {'length': 6.8, 'width': 860, 'qty': 12},  # ПБ 68-8-8п
        {'length': 6.8, 'width': 460, 'qty': 3},   # ПБ 68-4-8п
    ]
}

# Набор 5: Только узкие плиты (20×ПБ 56-3-8п)
TEST_CASE_5 = {
    'name': 'Набор 5: Только узкие (20×ПБ 56-3-8п)',
    'orders_2d': [
        {'length': 5.6, 'width': 320, 'qty': 20},  # ПБ 56-3-8п
    ]
}

# Все тестовые кейсы
TEST_CASES = [TEST_CASE_1, TEST_CASE_2, TEST_CASE_3, TEST_CASE_4, TEST_CASE_5]


# ==================== ФУНКЦИИ АНАЛИЗА ====================

def analyze_result(result: dict, orders_2d: list) -> dict:
    """
    Анализирует результат оптимизации и возвращает метрики.
    
    Returns:
        {
            'primary_cuts_count': int,          # Количество первичных резов
            'secondary_cuts_count': int,        # Количество вторичных резов
            'total_plates': int,                # Общее количество плит
            'waste_area_m2': float,             # Отходы (площадь)
            'unused_rests_count': int,          # Количество неиспользованных остатков
            'unused_rests_width_sum': float,    # Суммарная ширина неиспользованных остатков
        }
    """
    if not result:
        return {
            'primary_cuts_count': 0,
            'secondary_cuts_count': 0,
            'total_plates': 0,
            'waste_area_m2': 0.0,
            'unused_rests_count': 0,
            'unused_rests_width_sum': 0.0,
        }
    
    # Количество первичных резов
    primary_cuts_count = len(result.get('primary_cuts', []))
    
    # Количество вторичных резов
    secondary_cuts_count = len(result.get('secondary_cuts', []))
    
    # Общее количество плит
    total_plates = result.get('total_plates', 0)
    
    # Отходы (площадь)
    waste_area_m2 = 0.0
    for sec in result.get('secondary_cuts', []):
        waste_width_mm = sec.get('waste', 0)
        source_length = sec.get('source_lengths', [0])[0]  # берём первую длину
        if waste_width_mm > 0:
            waste_area_m2 += (waste_width_mm / 1000.0) * source_length * sec.get('qty', 0)
        
        length_waste_mm = sec.get('length_waste', 0)
        if length_waste_mm > 0:
            source_rest = sec.get('source', 0)
            waste_area_m2 += (length_waste_mm / 1000.0) * (source_rest / 1000.0) * sec.get('qty', 0)
    
    # Неиспользованные остатки
    # Считаем остатки, которые создали, но не использовали
    produced_rests = {}  # {(length, rest_width): qty}
    consumed_rests = {}  # {(length, rest_width): qty}
    
    # Произведено остатков
    for prim in result.get('primary_cuts', []):
        rest = prim.get('rest', 0)
        if rest > 0:
            for length in prim.get('lengths', []):
                key = (length, rest)
                produced_rests[key] = produced_rests.get(key, 0) + 1
    
    # Использовано остатков (через вторичные резы)
    for sec in result.get('secondary_cuts', []):
        source_rest = sec.get('source', 0)
        source_lengths = sec.get('source_lengths', [])
        qty = sec.get('qty', 0)
        for length in source_lengths:
            key = (length, source_rest)
            consumed_rests[key] = consumed_rests.get(key, 0) + 1
    
    # Неиспользованные остатки = произведено - использовано
    unused_rests_count = 0
    unused_rests_width_sum = 0.0
    for key, prod_qty in produced_rests.items():
        cons_qty = consumed_rests.get(key, 0)
        unused = prod_qty - cons_qty
        if unused > 0:
            unused_rests_count += unused
            unused_rests_width_sum += key[1] * unused  # ширина остатка × количество
    
    return {
        'primary_cuts_count': primary_cuts_count,
        'secondary_cuts_count': secondary_cuts_count,
        'total_plates': total_plates,
        'waste_area_m2': round(waste_area_m2, 3),
        'unused_rests_count': unused_rests_count,
        'unused_rests_width_sum': round(unused_rests_width_sum / 1000.0, 3),  # в метрах
    }


def run_experiment(test_case: dict, config: OptimizationConfig, config_name: str):
    """
    Запускает один эксперимент и выводит результаты.
    """
    print(f"\n{'='*70}")
    print(f"  РЕЖИМ: {config_name}")
    print(f"{'='*70}")
    
    orders_2d = test_case['orders_2d']
    
    # Запускаем оптимизацию
    result = optimize_with_cascading_longitudinal_cuts(
        orders_2d=orders_2d,
        opt_config=config
    )
    
    # Анализируем результаты
    metrics = analyze_result(result, orders_2d)
    
    # Выводим результаты
    print(f"\n📊 РЕЗУЛЬТАТЫ ({config_name}):")
    print(f"  • Первичных резов:           {metrics['primary_cuts_count']}")
    print(f"  • Вторичных резов:           {metrics['secondary_cuts_count']}")
    print(f"  • Всего плит использовано:   {metrics['total_plates']}")
    print(f"  • Отходы (площадь):          {metrics['waste_area_m2']} м²")
    print(f"  • Неиспользованных остатков: {metrics['unused_rests_count']} шт")
    print(f"  • Сумма ширин остатков:      {metrics['unused_rests_width_sum']} м")
    
    return metrics


def compare_results(old_metrics: dict, new_metrics: dict):
    """
    Сравнивает результаты OLD и NEW и выводит выводы.
    """
    print(f"\n{'='*70}")
    print(f"  📈 СРАВНЕНИЕ: OLD vs NEW")
    print(f"{'='*70}")
    
    # Сравниваем метрики
    print(f"\n┌─────────────────────────────────┬─────────┬─────────┬──────────┐")
    print(f"│ Метрика                         │   OLD   │   NEW   │ Разница  │")
    print(f"├─────────────────────────────────┼─────────┼─────────┼──────────┤")
    
    # Первичные резы
    old_prim = old_metrics['primary_cuts_count']
    new_prim = new_metrics['primary_cuts_count']
    diff_prim = new_prim - old_prim
    diff_str_prim = f"+{diff_prim}" if diff_prim > 0 else str(diff_prim)
    print(f"│ Первичных резов                 │ {old_prim:>7} │ {new_prim:>7} │ {diff_str_prim:>8} │")
    
    # Вторичные резы
    old_sec = old_metrics['secondary_cuts_count']
    new_sec = new_metrics['secondary_cuts_count']
    diff_sec = new_sec - old_sec
    diff_str_sec = f"+{diff_sec}" if diff_sec > 0 else str(diff_sec)
    print(f"│ Вторичных резов                 │ {old_sec:>7} │ {new_sec:>7} │ {diff_str_sec:>8} │")
    
    # Всего плит
    old_plates = old_metrics['total_plates']
    new_plates = new_metrics['total_plates']
    diff_plates = new_plates - old_plates
    diff_str_plates = f"+{diff_plates}" if diff_plates > 0 else str(diff_plates)
    print(f"│ Всего плит                      │ {old_plates:>7} │ {new_plates:>7} │ {diff_str_plates:>8} │")
    
    # Отходы
    old_waste = old_metrics['waste_area_m2']
    new_waste = new_metrics['waste_area_m2']
    diff_waste = new_waste - old_waste
    diff_str_waste = f"+{diff_waste:.2f}" if diff_waste > 0 else f"{diff_waste:.2f}"
    print(f"│ Отходы (м²)                     │ {old_waste:>7.2f} │ {new_waste:>7.2f} │ {diff_str_waste:>8} │")
    
    # Неиспользованные остатки
    old_unused = old_metrics['unused_rests_count']
    new_unused = new_metrics['unused_rests_count']
    diff_unused = new_unused - old_unused
    diff_str_unused = f"+{diff_unused}" if diff_unused > 0 else str(diff_unused)
    print(f"│ Неиспользованных остатков (шт)  │ {old_unused:>7} │ {new_unused:>7} │ {diff_str_unused:>8} │")
    
    # Сумма ширин остатков
    old_width = old_metrics['unused_rests_width_sum']
    new_width = new_metrics['unused_rests_width_sum']
    diff_width = new_width - old_width
    diff_str_width = f"+{diff_width:.2f}" if diff_width > 0 else f"{diff_width:.2f}"
    print(f"│ Сумма ширин остатков (м)        │ {old_width:>7.2f} │ {new_width:>7.2f} │ {diff_str_width:>8} │")
    
    print(f"└─────────────────────────────────┴─────────┴─────────┴──────────┘")
    
    # Выводы
    print(f"\n💡 ВЫВОДЫ:")
    
    if new_plates < old_plates:
        print(f"  ✅ NEW использует МЕНЬШЕ плит (на {old_plates - new_plates} шт) — экономия материала!")
    elif new_plates > old_plates:
        print(f"  ⚠️  NEW использует БОЛЬШЕ плит (на {new_plates - old_plates} шт) — расход материала вырос")
    else:
        print(f"  ➡️  Количество плит одинаковое")
    
    if new_sec < old_sec:
        print(f"  ✅ NEW делает МЕНЬШЕ вторичных резов (на {old_sec - new_sec} шт) — упрощение производства")
    elif new_sec > old_sec:
        print(f"  ⚠️  NEW делает БОЛЬШЕ вторичных резов (на {new_sec - old_sec} шт) — сложнее производство")
    else:
        print(f"  ➡️  Количество вторичных резов одинаковое")
    
    if new_unused > old_unused:
        print(f"  ✅ NEW оставляет БОЛЬШЕ крупных остатков (на {new_unused - old_unused} шт) — запас на будущее")
    elif new_unused < old_unused:
        print(f"  ⚠️  NEW оставляет МЕНЬШЕ остатков — максимальное использование")
    else:
        print(f"  ➡️  Количество остатков одинаковое")
    
    if new_waste < old_waste:
        print(f"  ✅ NEW даёт МЕНЬШЕ отходов (на {old_waste - new_waste:.2f} м²) — меньше мусора")
    elif new_waste > old_waste:
        print(f"  ⚠️  NEW даёт БОЛЬШЕ отходов (на {new_waste - old_waste:.2f} м²) — больше мусора")
    else:
        print(f"  ➡️  Отходы одинаковые")


# ==================== ГЛАВНАЯ ФУНКЦИЯ ====================

def main():
    """
    Запускает все эксперименты и сравнивает результаты.
    """
    print(f"\n{'#'*70}")
    print(f"  ЭКСПЕРИМЕНТЫ: СРАВНЕНИЕ OLD vs NEW МОДЕЛЕЙ ОПТИМИЗАЦИИ")
    print(f"{'#'*70}")
    print(f"\nOLD: unused_penalty=0.5, reuse_bonus=-500")
    print(f"NEW: unused_penalty=0.15, reuse_bonus=0")
    print(f"\nЗапускаем {len(TEST_CASES)} тестовых кейсов...\n")
    
    # Запускаем все тестовые кейсы
    for i, test_case in enumerate(TEST_CASES, start=1):
        print(f"\n{'#'*70}")
        print(f"  КЕЙС {i}: {test_case['name']}")
        print(f"{'#'*70}")
        
        # Показываем заказ
        print(f"\n📦 ЗАКАЗ:")
        for order in test_case['orders_2d']:
            print(f"  • {order['qty']}× плита {order['length']}м × {order['width']}мм")
        
        # Запускаем OLD конфигурацию
        old_metrics = run_experiment(test_case, OLD_CONFIG, "OLD")
        
        # Запускаем NEW конфигурацию
        new_metrics = run_experiment(test_case, NEW_CONFIG, "NEW")
        
        # Сравниваем результаты
        compare_results(old_metrics, new_metrics)
    
    print(f"\n{'#'*70}")
    print(f"  ✅ ВСЕ ЭКСПЕРИМЕНТЫ ЗАВЕРШЕНЫ!")
    print(f"{'#'*70}\n")


if __name__ == "__main__":
    main()

