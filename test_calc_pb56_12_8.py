#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для расчета себестоимости плиты ПБ 56-12-8п
"""

from cost_calculation import calculate_plate_cost
import core.config as cfg

# Рассчитываем себестоимость плиты ПБ 56-12-8п
plate_name = "ПБ 56-12-8п"
print("=" * 80)
print(f"РАСЧЕТ СЕБЕСТОИМОСТИ ПЛИТЫ: {plate_name}")
print("=" * 80)
print()

result = calculate_plate_cost(plate_name, cfg.PRICE_DB_PATH)

if result:
    print(f"📋 Параметры плиты:")
    params = result['parameters']
    print(f"   Длина: {params['length_dm']} дм = {params['length_m']} м")
    print(f"   Ширина: {params['width_dm']} дм = {params['width_m']} м")
    print(f"   Нагрузка: {params['load_code']}п")
    print(f"   Марка бетона: {params['concrete_grade']}")
    print()
    
    print(f"📊 Объем плиты: {result['volume_m3']:.4f} м³")
    print()
    
    print(f"💰 Компоненты себестоимости:")
    components = result['components']
    print(f"   Бетон:        {components['concrete']:>12,.2f} руб")
    print(f"   Армирование:  {components['reinforcement']:>12,.2f} руб")
    print(f"   Петли:        {components['loops']:>12,.2f} руб")
    print(f"   Изоформ:      {components['izoform']:>12,.2f} руб")
    print()
    
    print(f"📈 Прямые затраты:     {result['direct_cost']:>12,.2f} руб")
    print(f"📈 КЭФ:                {result['kef']:>12.2f}")
    print(f"📈 Накладные расходы:  {result['overhead_cost']:>12,.2f} руб")
    print(f"📈 Полная себестоимость: {result['full_cost_with_kef']:>12,.2f} руб")
    print()
    
    # Детальная разбивка
    breakdown = result['breakdown']
    print(f"🔍 Детальная разбивка:")
    print()
    
    print(f"   Бетон ({breakdown['concrete']['grade']}):")
    if 'cement_kg' in breakdown['concrete']:
        print(f"     Цемент:     {breakdown['concrete']['cement_kg']:>8.2f} кг")
        print(f"     Песок:      {breakdown['concrete']['sand_kg']:>8.2f} кг")
        print(f"     Щебень:     {breakdown['concrete']['gravel_kg']:>8.2f} кг")
        if 'polyplast_kg' in breakdown['concrete']:
            print(f"     Полипласт:  {breakdown['concrete']['polyplast_kg']:>8.2f} кг")
    print()
    
    print(f"   Армирование:")
    if 'wire_kg' in breakdown['reinforcement']:
        print(f"     Проволока:  {breakdown['reinforcement']['wire_kg']:>8.3f} кг")
        if 'wire_cost' in breakdown['reinforcement']:
            print(f"     Стоимость проволоки: {breakdown['reinforcement']['wire_cost']:>8.2f} руб")
        if 'cable_cost' in breakdown['reinforcement']:
            print(f"     Канат:     {breakdown['reinforcement']['cable_cost']:>8.2f} руб")
    print()
    
    if 'izoform_kg' in breakdown:
        print(f"   Изоформ:     {breakdown['izoform_kg']:>8.4f} кг")
    print()
    
    print("=" * 80)
    print("✅ РАСЧЕТ ЗАВЕРШЕН")
    print("=" * 80)
else:
    print("❌ Ошибка: не удалось рассчитать себестоимость плиты")
    print("   Проверьте, что данные загружены из Excel в БД")

