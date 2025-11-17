#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка цен в базе данных pb.db
"""
import sqlite3
import os
import sys

# Устанавливаем кодировку для вывода
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Путь к базе данных (в корневой папке проекта)
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pb.db')

print("=" * 70)
print("ПРОВЕРКА БАЗЫ ДАННЫХ pb.db")
print("=" * 70)
print(f"\nПуть к базе: {DB_PATH}")

if not os.path.exists(DB_PATH):
    print("\nОШИБКА: База данных не найдена!")
    exit(1)

print("OK: База данных найдена\n")

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

# Проверяем количество записей
cur.execute("SELECT COUNT(*) FROM prices")
count = cur.fetchone()[0]
print(f"Всего записей в таблице prices: {count}")

# Ищем цену для плиты 3.6м (36 дм) с нагрузкой 8
length_m = 3.6
load_code = 8
length_dm = int(round(length_m * 10))

print(f"\nИЩУ ЦЕНУ ДЛЯ ПЛИТЫ:")
print(f"   Длина: {length_m} м = {length_dm} дм")
print(f"   Класс нагрузки: {load_code}")

# Точный поиск
cur.execute('SELECT price FROM prices WHERE length_dm=? AND load_code=?', (length_dm, load_code))
row = cur.fetchone()

if row:
    price = float(row[0])
    print(f"\nНАЙДЕНА ТОЧНАЯ ЦЕНА: {price:,.2f} руб")
    print(f"   (форматировано: {price:,.2f} руб)")
else:
    print(f"\nВНИМАНИЕ: Точная цена не найдена для length_dm={length_dm}, load_code={load_code}")
    print("   Ищу ближайшую цену (с допуском ±1 дм)...")
    
    # Поиск с допуском
    cur.execute('SELECT length_dm, load_code, price FROM prices WHERE ABS(length_dm-?)<=1 AND load_code=? ORDER BY ABS(length_dm-?) LIMIT 1', 
                (length_dm, load_code, length_dm))
    row = cur.fetchone()
    
    if row:
        found_length_dm, found_load_code, price = row
        print(f"\nНАЙДЕНА БЛИЖАЙШАЯ ЦЕНА:")
        print(f"   length_dm={found_length_dm} (искали {length_dm})")
        print(f"   load_code={found_load_code}")
        print(f"   price={price:,.2f} руб")
    else:
        print("\nОШИБКА: Цена не найдена даже с допуском!")

# Показываем ВСЕ цены для длины около 36 дм (±3 дм)
print(f"\nВСЕ ЦЕНЫ ДЛЯ ДЛИНЫ ОКОЛО {length_dm} дм (±3 дм):")
print("-" * 70)
cur.execute('SELECT length_dm, load_code, price FROM prices WHERE ABS(length_dm-?)<=3 ORDER BY length_dm, load_code', (length_dm,))
rows = cur.fetchall()

if rows:
    print(f"{'Длина (дм)':<12} {'Нагрузка':<10} {'Цена (руб)':<20}")
    print("-" * 70)
    for r in rows:
        length_dm_val, load_code_val, price_val = r
        print(f"{length_dm_val:<12} {load_code_val:<10} {price_val:>15,.2f}")
else:
    print("   (нет записей)")

# Показываем первые 15 записей из базы
print(f"\nПЕРВЫЕ 15 ЗАПИСЕЙ ИЗ БАЗЫ ДАННЫХ:")
print("-" * 70)
cur.execute('SELECT length_dm, load_code, price FROM prices ORDER BY length_dm, load_code LIMIT 15')
rows = cur.fetchall()

if rows:
    print(f"{'Длина (дм)':<12} {'Нагрузка':<10} {'Цена (руб)':<20}")
    print("-" * 70)
    for r in rows:
        length_dm_val, load_code_val, price_val = r
        print(f"{length_dm_val:<12} {load_code_val:<10} {price_val:>15,.2f}")
else:
    print("   (база данных пуста)")

conn.close()

print("\n" + "=" * 70)
print("Проверка завершена!")
print("=" * 70)

