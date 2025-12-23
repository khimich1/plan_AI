#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка данных в БД для плиты ПБ 56-12-8п
"""

import sqlite3
import core.config_and_data as cfg

db_path = cfg.PRICE_DB_PATH

print("=" * 80)
print("ПРОВЕРКА ДАННЫХ В БД ДЛЯ ПЛИТЫ ПБ 56-12-8п")
print("=" * 80)
print()

conn = sqlite3.connect(db_path)
try:
    cur = conn.cursor()
    
    length_dm = 56
    width_dm = 12
    load_code = 8
    
    # Проверяем объем
    print("📊 ОБЪЕМ ПЛИТЫ:")
    cur.execute("""
        SELECT volume_m3 FROM plate_volumes
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    row = cur.fetchone()
    if row:
        print(f"   Объем из БД: {row[0]:.4f} м³")
    else:
        print("   ❌ Объем не найден в БД")
    print()
    
    # Проверяем стоимость бетона
    print("🏗️ СТОИМОСТЬ БЕТОНА:")
    cur.execute("""
        SELECT concrete_cost FROM concrete_costs
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    row = cur.fetchone()
    if row:
        print(f"   Стоимость бетона из БД: {row[0]:,.2f} руб")
    else:
        print("   ❌ Стоимость бетона не найдена в БД")
    print()
    
    # Проверяем армирование
    print("🔩 АРМИРОВАНИЕ:")
    cur.execute("""
        SELECT wire_kg, cable_cost FROM reinforcement_costs
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    row = cur.fetchone()
    if row:
        wire_kg = row[0] if row[0] else 0
        cable_cost = row[1] if row[1] else 0
        print(f"   Проволока: {wire_kg:.3f} кг")
        print(f"   Канат: {cable_cost:,.2f} руб")
        
        # Рассчитываем стоимость проволоки
        cur.execute("SELECT value FROM cost_constants WHERE key='wire_price_per_kg'")
        wire_price_row = cur.fetchone()
        wire_price = wire_price_row[0] if wire_price_row else 80.0
        wire_cost = wire_kg * wire_price
        print(f"   Стоимость проволоки: {wire_cost:,.2f} руб (при цене {wire_price} руб/кг)")
        print(f"   ИТОГО армирование: {wire_cost + cable_cost:,.2f} руб")
    else:
        print("   ❌ Данные армирования не найдены в БД")
    print()
    
    # Проверяем изоформ
    print("💧 ИЗОФОРМ:")
    cur.execute("""
        SELECT izoform_kg, izoform_cost FROM izoform_costs
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    row = cur.fetchone()
    if row:
        izoform_kg = row[0] if row[0] else 0
        izoform_cost = row[1] if row[1] else 0
        print(f"   Количество: {izoform_kg:.4f} кг")
        print(f"   Стоимость: {izoform_cost:.2f} руб")
    else:
        print("   ❌ Данные изоформа не найдены в БД")
    print()
    
    # Проверяем КЭФ
    print("📈 КЭФ:")
    cur.execute("""
        SELECT kef FROM plate_kef_values
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    row = cur.fetchone()
    if row:
        print(f"   КЭФ для плиты: {row[0]:.2f}")
    else:
        cur.execute("SELECT value FROM cost_constants WHERE key='kef'")
        kef_row = cur.fetchone()
        kef_default = kef_row[0] if kef_row else 1.25
        print(f"   КЭФ (дефолт): {kef_default:.2f}")
    print()
    
    # Проверяем общую себестоимость из Excel (для сравнения)
    print("📋 ОБЩАЯ СЕБЕСТОИМОСТЬ ИЗ EXCEL (для проверки):")
    cur.execute("""
        SELECT total_cost FROM excel_total_costs
        WHERE length_dm=? AND width_dm=? AND load_code=?
    """, (length_dm, width_dm, load_code))
    row = cur.fetchone()
    if row:
        print(f"   Итого из Excel: {row[0]:,.2f} руб")
    else:
        print("   ⚠️ Общая себестоимость из Excel не найдена")
    print()
    
    print("=" * 80)
    
finally:
    conn.close()

