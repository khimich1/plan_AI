#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Анализ логики расчета армирования
"""

import sqlite3
import core.config_and_data as cfg

db_path = cfg.PRICE_DB_PATH

print("=" * 80)
print("АНАЛИЗ ЛОГИКИ РАСЧЕТА АРМИРОВАНИЯ")
print("=" * 80)
print()

conn = sqlite3.connect(db_path)
try:
    cur = conn.cursor()
    
    length_dm = 56
    width_dm = 12
    load_code = 8
    
    # Получаем данные из БД
    cur.execute("""
        SELECT wire_kg, cable_cost FROM reinforcement_costs
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    row = cur.fetchone()
    
    if not row:
        print("❌ Данные не найдены")
        exit(1)
    
    wire_kg = row[0] if row[0] else 0
    cable_cost_from_db = row[1] if row[1] else 0
    
    # Получаем цену проволоки
    cur.execute("SELECT value FROM cost_constants WHERE key='wire_price_per_kg'")
    wire_price_row = cur.fetchone()
    wire_price = wire_price_row[0] if wire_price_row else 80.0
    
    print("📊 ДАННЫЕ ИЗ БД:")
    print(f"   Проволока: {wire_kg:.3f} кг")
    print(f"   Канат (колонка 6): {cable_cost_from_db:,.2f} руб")
    print()
    
    print("🔍 РАСЧЕТ:")
    print()
    
    # Вариант 1: Канат - это только стоимость каната
    wire_cost_v1 = wire_kg * wire_price
    total_v1 = wire_cost_v1 + cable_cost_from_db
    print(f"ВАРИАНТ 1: Канат = только стоимость каната")
    print(f"   Проволока: {wire_kg:.3f} кг × {wire_price} руб/кг = {wire_cost_v1:,.2f} руб")
    print(f"   Канат: {cable_cost_from_db:,.2f} руб")
    print(f"   ИТОГО: {total_v1:,.2f} руб")
    print()
    
    # Вариант 2: Канат - это уже полная стоимость армирования
    print(f"ВАРИАНТ 2: Канат = полная стоимость армирования (проволока + канат)")
    print(f"   ИТОГО: {cable_cost_from_db:,.2f} руб")
    print()
    
    print("=" * 80)
    print("СРАВНЕНИЕ С ДАННЫМИ ИЗ СКРИНШОТОВ:")
    print("=" * 80)
    print()
    print("Из скриншотов:")
    print("   Желтая колонка: 8,104.5 руб")
    print("   Колонка 19 (общая сумма): 8,104.46 руб")
    print()
    
    print("Анализ:")
    print()
    
    if abs(cable_cost_from_db - 8104.5) < 1:
        print("✅ Колонка 6 (канат) = 8,104.5 руб совпадает с желтой колонкой!")
        print("   → Значит, в колонке 6 УЖЕ полная стоимость армирования")
        print("   → НЕ нужно добавлять стоимость проволоки!")
    elif abs(total_v1 - 8104.5) < 1:
        print("✅ Наш расчет (проволока + канат) = 8,104.5 руб совпадает!")
        print("   → Значит, нужно добавлять стоимость проволоки")
    else:
        print("⚠️  Не совпадает ни один вариант")
        print(f"   Колонка 6: {cable_cost_from_db:,.2f} руб")
        print(f"   Наш расчет: {total_v1:,.2f} руб")
        print(f"   Желтая колонка: 8,104.5 руб")
        print()
        print("   Возможные причины:")
        print("   1. В Excel используется другая цена проволоки")
        print("   2. В Excel другая логика расчета")
        print("   3. Желтая колонка - это другая колонка (не колонка 6)")
    
    print()
    
finally:
    conn.close()

