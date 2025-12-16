#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пример использования функции расчета себестоимости плит
"""

from cost_calculation import calculate_plate_cost, init_default_constants
import core.config_and_data as cfg

# Инициализируем БД с константами (один раз, при первом запуске)
print("Инициализация БД с константами...")
init_default_constants(cfg.PRICE_DB_PATH)
print("✅ БД инициализирована\n")

# Примеры расчета себестоимости
examples = [
    "ПБ 17-12-6",
    "ПБ 20-12-8",
    "ПБ 30-12-10",
]

print("=" * 80)
print("РАСЧЕТ СЕБЕСТОИМОСТИ ПЛИТ")
print("=" * 80)
print()

for plate_name in examples:
    result = calculate_plate_cost(plate_name, cfg.PRICE_DB_PATH)
    
    if result:
        print(f"📋 Плита: {result['plate_name']}")
        print(f"   Параметры:")
        params = result['parameters']
        print(f"     - Длина: {params['length_m']:.1f} м ({params['length_dm']} дм)")
        print(f"     - Ширина: {params['width_m']:.1f} м ({params['width_dm']} дм)")
        print(f"     - Нагрузка: {params['load_code']}п")
        print(f"     - Марка бетона: {params['concrete_grade']}")
        print(f"     - Объем: {result['volume_m3']:.4f} м³")
        
        print(f"\n   Компоненты себестоимости:")
        components = result['components']
        for component, cost in components.items():
            print(f"     - {component.capitalize()}: {cost:,.2f} руб")
        
        print(f"\n   💰 ИТОГО: {result['total_cost']:,.2f} руб")
        
        print(f"\n   Детальная разбивка:")
        breakdown = result['breakdown']
        
        concrete = breakdown['concrete']
        print(f"     Бетон ({concrete['grade']}):")
        print(f"       - Цемент: {concrete['cement_kg']:.3f} кг")
        print(f"       - Песок: {concrete['sand_m3']:.4f} м³")
        print(f"       - Щебень: {concrete['gravel_m3']:.4f} м³")
        
        reinforcement = breakdown['reinforcement']
        print(f"     Армирование (нагрузка {reinforcement['load_code']}п):")
        print(f"       - Проволока: {reinforcement['wire_kg']:.3f} кг")
        print(f"       - Канат: {reinforcement['cable_cost']:.2f} руб")
        
        print()
        print("-" * 80)
        print()
    else:
        print(f"❌ Ошибка: не удалось рассчитать себестоимость для '{plate_name}'")
        print()

print("=" * 80)
print("✅ Расчет завершен")
print("=" * 80)

