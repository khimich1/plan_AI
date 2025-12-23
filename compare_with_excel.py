#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сравнение расчета с данными из Excel для ПБ 56-12-8п
"""

from cost_calculation import calculate_plate_cost
import core.config_and_data as cfg
import sqlite3

plate_name = "ПБ 56-12-8п"

print("=" * 80)
print("СРАВНЕНИЕ РАСЧЕТА С ДАННЫМИ ИЗ EXCEL")
print(f"Плита: {plate_name}")
print("=" * 80)
print()

# Получаем расчет
result = calculate_plate_cost(plate_name, cfg.PRICE_DB_PATH)

if not result:
    print("❌ Не удалось рассчитать")
    exit(1)

# Получаем данные из Excel для сравнения
conn = sqlite3.connect(cfg.PRICE_DB_PATH)
try:
    cur = conn.cursor()
    
    length_dm = 56
    width_dm = 12
    load_code = 8
    
    # Получаем общую себестоимость из Excel
    cur.execute("""
        SELECT total_cost FROM excel_total_costs
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    excel_total = cur.fetchone()
    
    print("📊 СРАВНЕНИЕ КОМПОНЕНТОВ:")
    print()
    
    print(f"1. ОБЪЕМ:")
    print(f"   Наш расчет: {result['volume_m3']:.4f} м³")
    print(f"   Из Excel:   0.9240 м³")
    print(f"   ✅ Совпадает: {abs(result['volume_m3'] - 0.9240) < 0.0001}")
    print()
    
    print(f"2. БЕТОН:")
    print(f"   Наш расчет: {result['components']['concrete']:,.2f} руб")
    print(f"   Из Excel:   5,646.52 руб (колонка 13)")
    print(f"   ✅ Совпадает: {abs(result['components']['concrete'] - 5646.52) < 0.01}")
    print()
    
    print(f"3. АРМИРОВАНИЕ:")
    print(f"   Наш расчет: {result['components']['reinforcement']:,.2f} руб")
    breakdown = result['breakdown']
    if 'wire_cost' in breakdown['reinforcement']:
        print(f"     - Проволока: {breakdown['reinforcement']['wire_cost']:,.2f} руб")
    if 'cable_cost' in breakdown['reinforcement']:
        print(f"     - Канат: {breakdown['reinforcement']['cable_cost']:,.2f} руб")
    print(f"   Из Excel:   8,104.5 руб (желтая колонка)")
    print(f"   ⚠️  Разница: {abs(result['components']['reinforcement'] - 8104.5):,.2f} руб")
    print()
    
    print(f"4. ПЕТЛИ:")
    print(f"   Наш расчет: {result['components']['loops']:,.2f} руб")
    print(f"   Из Excel:   572 руб (стандартно)")
    print(f"   ✅ Совпадает: {abs(result['components']['loops'] - 572) < 0.01}")
    print()
    
    print(f"5. ИЗОФОРМ:")
    print(f"   Наш расчет: {result['components']['izoform']:,.2f} руб")
    print(f"   Из Excel:   38.30 руб")
    print(f"   ✅ Совпадает: {abs(result['components']['izoform'] - 38.30) < 0.01}")
    print()
    
    print(f"6. ПРЯМЫЕ ЗАТРАТЫ:")
    print(f"   Наш расчет: {result['direct_cost']:,.2f} руб")
    if excel_total:
        print(f"   Из Excel (колонка 19): {excel_total[0]:,.2f} руб")
        diff = abs(result['direct_cost'] - excel_total[0])
        print(f"   ⚠️  Разница: {diff:,.2f} руб ({diff/result['direct_cost']*100:.1f}%)")
    print()
    
    print(f"7. ПОЛНАЯ СЕБЕСТОИМОСТЬ С КЭФ:")
    print(f"   Наш расчет: {result['full_cost_with_kef']:,.2f} руб")
    print(f"   КЭФ: {result['kef']:.2f}")
    print()
    
    # Проверяем, что могло быть в желтой колонке (8104.5)
    print("=" * 80)
    print("АНАЛИЗ РАСХОЖДЕНИЙ:")
    print("=" * 80)
    print()
    print("Из скриншотов видно:")
    print("  - Желтая колонка (8104.5 руб) - возможно, это стоимость армирования")
    print("  - Колонка 19 (8,104.46 руб) - возможно, это прямые затраты БЕЗ чего-то")
    print()
    print("Наш расчет армирования:")
    print(f"  Проволока: {breakdown['reinforcement'].get('wire_cost', 0):,.2f} руб")
    print(f"  Канат: {breakdown['reinforcement'].get('cable_cost', 0):,.2f} руб")
    print(f"  ИТОГО: {result['components']['reinforcement']:,.2f} руб")
    print()
    print("⚠️  ВОЗМОЖНАЯ ПРОБЛЕМА:")
    print("   В Excel желтая колонка (8104.5) может быть:")
    print("   1. Только стоимость каната (но у нас 1990.63)")
    print("   2. Стоимость армирования с другими данными")
    print("   3. Другая колонка (не армирование)")
    print()
    
finally:
    conn.close()

