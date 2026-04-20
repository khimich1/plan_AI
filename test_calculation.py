#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Тест расчета себестоимости"""

from core.cost_calculation import calculate_plate_cost
import core.config as cfg

# Тестируем плиту ПБ 17-12-6
result = calculate_plate_cost("ПБ 17-12-6", cfg.PRICE_DB_PATH)

if result:
    print(f"Плита: {result['plate_name']}")
    print(f"Объем: {result['volume_m3']:.4f} м³")
    print(f"Компоненты:")
    for comp, cost in result['components'].items():
        print(f"  {comp}: {cost:,.2f} руб")
    print(f"ИТОГО: {result['total_cost']:,.2f} руб")
    
    # Проверяем с Excel
    import sqlite3
    conn = sqlite3.connect(cfg.PRICE_DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT total_cost FROM excel_total_costs
        WHERE length_dm=17 AND width_dm=12 AND load_code=6
    """)
    row = cur.fetchone()
    if row:
        excel_cost = row[0]
        print(f"\nСебестоимость из Excel: {excel_cost:,.2f} руб")
        diff = abs(result['total_cost'] - excel_cost)
        print(f"Разница: {diff:,.2f} руб")
    conn.close()
else:
    print("Ошибка расчета")

