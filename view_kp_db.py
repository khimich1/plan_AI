#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для просмотра содержимого базы данных plita.db
"""

import sys
import sqlite3
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent
db_path = PROJECT_ROOT / 'plita.db'

if not db_path.exists():
    print("❌ База данных plita.db не найдена!")
    print(f"   Путь: {db_path}")
    sys.exit(1)

print("=" * 70)
print("📊 ПРОСМОТР БАЗЫ ДАННЫХ plita.db")
print("=" * 70)

conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

# Показываем все КП
cur.execute('''
    SELECT 
        ko.kp_id,
        ko.creation_date,
        ko.customer_name,
        ko.manager_name,
        ko.discount_percent,
        ko.subtotal,
        ko.vat_amount,
        ko.total_amount,
        ko.execution_terms,
        km.status
    FROM KP_offers ko
    LEFT JOIN kp_meta km ON ko.kp_id = km.kp_id
    ORDER BY ko.kp_id DESC
''')

kps = cur.fetchall()

if not kps:
    print("\n❌ В базе данных нет КП")
else:
    print(f"\n✅ Найдено КП: {len(kps)}\n")
    
    for kp in kps:
        print(f"{'=' * 70}")
        print(f"📋 КП № {kp['kp_id']}")
        print(f"{'=' * 70}")
        print(f"📅 Дата создания: {kp['creation_date']}")
        print(f"👤 Клиент: {kp['customer_name']}")
        print(f"👨‍💼 Менеджер: {kp['manager_name']}")
        print(f"💰 Скидка: {kp['discount_percent']}%")
        print(f"💵 Сумма без НДС: {kp['subtotal']:,.2f} ₽")
        print(f"📊 НДС (20%): {kp['vat_amount']:,.2f} ₽")
        print(f"💎 ИТОГО: {kp['total_amount']:,.2f} ₽")
        print(f"⏰ Сроки: {kp['execution_terms'] or 'не указаны'}")
        print(f"📌 Статус: {kp['status'] or 'не указан'}")
        
        # Показываем плиты
        cur.execute('''
            SELECT 
                position_number,
                plate_name,
                qty,
                length_m,
                width_m,
                load_class,
                unit_weight,
                total_weight,
                discounted_price
            FROM kp_plates
            WHERE kp_id = ?
            ORDER BY position_number
        ''', (kp['kp_id'],))
        
        plates = cur.fetchall()
        
        print(f"\n📦 Плиты в заказе ({len(plates)} позиций):")
        for plate in plates:
            print(f"  {plate['position_number']}. {plate['plate_name']}")
            print(f"     • Количество: {plate['qty']} шт")
            print(f"     • Размеры: {plate['length_m']}м × {plate['width_m']}м")
            print(f"     • Нагрузка: {plate['load_class']} кг/м²")
            print(f"     • Вес: {plate['total_weight']:.2f} кг ({plate['unit_weight']:.2f} кг/шт)")
            print(f"     • Цена: {plate['discounted_price']:,.2f} ₽/шт")
        
        # Проверяем наличие файла
        cur.execute('SELECT file_path FROM kp_files WHERE kp_id = ?', (kp['kp_id'],))
        file_row = cur.fetchone()
        if file_row and file_row['file_path']:
            print(f"\n📁 Файл XLSX: {file_row['file_path']}")
        
        print()

conn.close()

print("=" * 70)
print("✅ Просмотр завершён")
print("=" * 70)
