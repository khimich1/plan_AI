#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка Excel файла и перезагрузка базы данных
"""
import os
import sys
import pandas as pd

# Устанавливаем кодировку для вывода
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ищем Excel файл с ценами
bank_dir = os.path.join(BASE_DIR, 'банк знаний')
PRICE_XLSX_PATH = None

if os.path.exists(bank_dir):
    for name in os.listdir(bank_dir):
        # Пропускаем временные файлы Excel (начинаются с ~$)
        if name.startswith('~$'):
            continue
        if name.lower().endswith('.xlsx') and 'нов' in name.lower() and 'цен' in name.lower():
            PRICE_XLSX_PATH = os.path.join(bank_dir, name)
            break

if not PRICE_XLSX_PATH or not os.path.exists(PRICE_XLSX_PATH):
    print("\nОШИБКА: Excel файл с ценами не найден!")
    print(f"Искал в: {bank_dir}")
    exit(1)

print("=" * 70)
print("ПРОВЕРКА EXCEL ФАЙЛА И ПЕРЕЗАГРУЗКА БАЗЫ ДАННЫХ")
print("=" * 70)
print(f"\nНайден Excel файл: {PRICE_XLSX_PATH}")

print("OK: Excel файл найден\n")

# Читаем Excel
try:
    all_sheets = pd.read_excel(PRICE_XLSX_PATH, sheet_name=None)
    print(f"Найдено листов в Excel: {len(all_sheets)}")
    print(f"Названия листов: {list(all_sheets.keys())}\n")
    
    # Ищем лист с датой 24.06.2024 или берем первый
    preferred_sheet = None
    for sheet_name in all_sheets.keys():
        if '24.06.2024' in str(sheet_name):
            preferred_sheet = sheet_name
            break
    
    if not preferred_sheet:
        preferred_sheet = list(all_sheets.keys())[0]
    
    print(f"Используем лист: {preferred_sheet}\n")
    
    df = all_sheets[preferred_sheet]
    
    # Ищем колонку "Наименование"
    name_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'наименование' or 'наимен' in str(c).strip().lower():
            name_col = c
            break
    
    if not name_col:
        print("ОШИБКА: Не найдена колонка 'Наименование'")
        print(f"Доступные колонки: {list(df.columns)}")
        exit(1)
    
    print(f"Колонка 'Наименование': {name_col}")
    
    # Ищем колонки с нагрузками
    headers = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl == '6 нагрузка':
            headers[6] = c
        elif cl == '8 нагрузка':
            headers[8] = c
        elif cl == '10 нагрузка':
            headers[10] = c
        elif cl == '12 нагрузка':
            headers[12] = c
    
    if not headers:
        # Пробуем найти в первой строке
        if len(df) > 0:
            first_row = df.iloc[0]
            for c in df.columns:
                val = str(first_row.get(c, '')).strip().lower()
                if 'нагруз' in val:
                    if '6' in val:
                        headers[6] = c
                    if '8' in val:
                        headers[8] = c
                    if '10' in val:
                        headers[10] = c
                    if '12' in val:
                        headers[12] = c
    
    print(f"Колонки с нагрузками: {headers}\n")
    
    # Ищем плиту 36 дм (3.6 м)
    import re
    target_length_dm = 36
    target_load = 8
    
    print(f"ИЩУ ПЛИТУ: длина {target_length_dm} дм, нагрузка {target_load}")
    print("-" * 70)
    
    found = False
    for idx, row in df.iterrows():
        name = str(row.get(name_col, '')).strip()
        if not name:
            continue
        
        # Ищем паттерн "36-0,3" или "36 - 0,3"
        m = re.search(r'(\d+)\s*-\s*(\d+)', name)
        if m:
            length_dm = int(m.group(1))
            if length_dm == target_length_dm:
                print(f"\nНАЙДЕНА ПЛИТА: {name}")
                print(f"  Индекс строки: {idx}")
                
                # Показываем все цены для этой плиты
                if headers:
                    for load_code, col in headers.items():
                        val = row.get(col)
                        if pd.notna(val):
                            try:
                                price = float(str(val).replace(' ', '').replace(',', '.'))
                                marker = " <-- ЭТО НАША ЦЕНА" if load_code == target_load else ""
                                print(f"  Нагрузка {load_code}: {price:,.2f} руб{marker}")
                            except Exception as e:
                                print(f"  Нагрузка {load_code}: ОШИБКА чтения ({e})")
                else:
                    print("  ОШИБКА: Не найдены колонки с нагрузками")
                
                found = True
                break
    
    if not found:
        print(f"\nВНИМАНИЕ: Плита длиной {target_length_dm} дм не найдена в Excel!")
        print("\nПоказываю первые 10 строк для проверки:")
        print("-" * 70)
        for idx in range(min(10, len(df))):
            name = str(df.iloc[idx].get(name_col, '')).strip()
            if name:
                print(f"  {idx}: {name}")
    
    # Теперь перезагружаем базу данных
    print("\n" + "=" * 70)
    print("ПЕРЕЗАГРУЗКА БАЗЫ ДАННЫХ")
    print("=" * 70)
    
    from price_db import import_from_xlsx
    from config_and_data import PRICE_DB_PATH
    
    count = import_from_xlsx(PRICE_XLSX_PATH, PRICE_DB_PATH, preferred_sheet)
    print(f"\nЗагружено записей в базу данных: {count}")
    
    # Проверяем, что загрузилось
    import sqlite3
    conn = sqlite3.connect(PRICE_DB_PATH)
    cur = conn.cursor()
    cur.execute('SELECT price FROM prices WHERE length_dm=? AND load_code=?', (target_length_dm, target_load))
    row = cur.fetchone()
    if row:
        print(f"Проверка: В базе теперь цена {row[0]:,.2f} руб для {target_length_dm} дм, нагрузка {target_load}")
    conn.close()
    
except Exception as e:
    print(f"\nОШИБКА: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 70)
print("Проверка завершена!")
print("=" * 70)

