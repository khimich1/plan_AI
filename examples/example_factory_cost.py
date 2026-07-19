#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пример использования модуля factory_cost

Демонстрирует:
- Получение себестоимости по параметрам
- Получение себестоимости по названию
- Детальную разбивку затрат
- Работу с заказом
"""

import os
import sys

# Добавляем корень проекта в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from factory_cost import get_cost_by_plate_name, get_cost_by_params
from factory_cost.cost_engine import get_cost_breakdown, get_all_available_plates
from core.config_and_data import set_plate_lists_from_text


def example_1_simple():
    """Пример 1: Простой запрос себестоимости"""
    print("\n" + "="*70)
    print("ПРИМЕР 1: Простой запрос себестоимости")
    print("="*70)
    
    # Получаем себестоимость плиты 7.1м × 1.2м
    cost = get_cost_by_params(length_m=7.1, width_m=1.2)
    
    if cost:
        print(f"\nПлита: {cost['plate_name']}")
        print(f"Размеры: {cost['length_dm']/10:.1f}м × {cost['width_dm']/10:.1f}м")
        print(f"Нагрузка: {cost['load_code']}п")
        print(f"Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
        print(f"Прямые затраты: {cost['direct_cost']:.2f} руб")
        print(f"Объём бетона: {cost['volume_m3']:.3f} м³" if cost['volume_m3'] else "")
    else:
        print("\n❌ Плита не найдена в базе себестоимости")


def example_2_by_name():
    """Пример 2: Запрос по названию плиты"""
    print("\n" + "="*70)
    print("ПРИМЕР 2: Запрос по названию плиты")
    print("="*70)
    
    plate_name = "Плиты ПБ 71-12-8п"
    cost = get_cost_by_plate_name(plate_name)
    
    if cost:
        print(f"\nПлита: {plate_name}")
        print(f"Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
        
        print("\nКомпоненты:")
        for comp, value in cost['components'].items():
            print(f"  {comp:15s}: {value:10.2f} руб")
    else:
        print(f"\n❌ Плита '{plate_name}' не найдена")


def example_3_breakdown():
    """Пример 3: Детальная разбивка затрат"""
    print("\n" + "="*70)
    print("ПРИМЕР 3: Детальная разбивка затрат")
    print("="*70)
    
    plate_name = "Плиты ПБ 71-12-10п"
    breakdown = get_cost_breakdown(plate_name)
    
    if breakdown:
        print(f"\nПлита: {plate_name}")
        print(f"КЭФ: {breakdown['kef']}")
        print(f"\nПрямые затраты:    {breakdown['direct_cost']:10.2f} руб")
        print(f"Накладные (КЭФ):   {breakdown['overhead_cost']:10.2f} руб")
        print(f"{'─'*40}")
        print(f"ИТОГО:             {breakdown['full_cost_with_kef']:10.2f} руб")
        
        print("\n\nСтруктура прямых затрат:")
        print(f"{'─'*40}")
        for comp, data in breakdown['breakdown'].items():
            print(f"{comp:15s}: {data['value']:10.2f} руб ({data['percentage']:5.1f}%)")
    else:
        print(f"\n❌ Плита '{plate_name}' не найдена")


def example_4_order():
    """Пример 4: Расчёт себестоимости для заказа"""
    print("\n" + "="*70)
    print("ПРИМЕР 4: Расчёт себестоимости для заказа")
    print("="*70)
    
    # Текст заказа
    order_text = """
    Плиты ПБ 71-12-10п - 5 шт
    Плиты ПБ 63-12-8п - 3 шт
    Плиты ПБ 54-12-8п - 2 шт
    """
    
    print(f"\nЗаказ:{order_text}")
    
    # Парсим заказ (загружает PLATE_LOAD_DETAILS для определения нагрузок)
    set_plate_lists_from_text(order_text)
    
    # Расчёт себестоимости
    orders = [
        (7.1, 1.2, 5),
        (6.3, 1.2, 3),
        (5.4, 1.2, 2),
    ]
    
    print("\n" + "─"*70)
    print(f"{'Плита':<25} {'Цена/шт':>12} {'Кол-во':>8} {'Сумма':>12}")
    print("─"*70)
    
    total_cost = 0.0
    
    for length_m, width_m, qty in orders:
        cost = get_cost_by_params(length_m, width_m)
        
        if cost:
            unit_cost = cost['full_cost_with_kef']
            line_cost = unit_cost * qty
            total_cost += line_cost
            
            print(f"{cost['plate_name']:<25} {unit_cost:12.2f} {qty:8d} {line_cost:12.2f}")
        else:
            print(f"ПБ {int(length_m*10)}-{int(width_m*10):<17} {'НЕ НАЙДЕНО':>12} {qty:8d} {0.0:12.2f}")
    
    print("─"*70)
    print(f"{'ИТОГО:':<25} {'':<12} {'':<8} {total_cost:12.2f}")
    print("─"*70)


def example_5_analysis():
    """Пример 5: Анализ структуры затрат"""
    print("\n" + "="*70)
    print("ПРИМЕР 5: Анализ структуры затрат (первые 5 плит)")
    print("="*70)
    
    # Получаем все плиты
    plates = get_all_available_plates()
    
    print(f"\nВсего плит в базе: {len(plates)}\n")
    
    # Анализируем первые 5
    for plate in plates[:5]:
        breakdown = get_cost_breakdown(plate['plate_name'])
        if breakdown:
            reinforcement_pct = breakdown['breakdown']['reinforcement']['percentage']
            concrete_pct = breakdown['breakdown']['concrete']['percentage']
            
            print(f"{plate['plate_name']:<25} "
                  f"Арм: {reinforcement_pct:5.1f}%, "
                  f"Бетон: {concrete_pct:5.1f}%")


def main():
    """Запуск всех примеров"""
    print("\n" + "="*70)
    print("ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ МОДУЛЯ factory_cost")
    print("="*70)
    
    try:
        example_1_simple()
        example_2_by_name()
        example_3_breakdown()
        example_4_order()
        example_5_analysis()
        
        print("\n" + "="*70)
        print("✅ Все примеры выполнены успешно!")
        print("="*70)
        print("\nДля импорта данных из Excel:")
        print("  python scripts/import_factory_costs.py")
        print("\nДля валидации данных:")
        print("  python scripts/validate_factory_costs.py")
        print()
        
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()

