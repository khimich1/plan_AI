#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка данных из Excel для плиты ПБ 56-12-8п
"""

import pandas as pd
import os
import core.config_and_data as cfg

EXCEL_PATH = os.path.join(cfg.BASE_DIR, "банк знаний", "Расчет новых цен на ПБ 10.09.2025 (1).xls")

print("=" * 80)
print("ПРОВЕРКА ДАННЫХ ИЗ EXCEL ДЛЯ ПЛИТЫ ПБ 56-12-8п")
print("=" * 80)
print()

try:
    df = pd.read_excel(EXCEL_PATH, sheet_name="Нов Серия для произв", engine='xlrd')
except:
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name="Нов Серия для произв", engine='openpyxl')
    except Exception as e:
        print(f"❌ Ошибка открытия Excel: {e}")
        exit(1)

# Находим строку с плитой ПБ 56-12-8п
plate_row = None
for idx, row in df.iterrows():
    plate_name = str(row.iloc[0])
    if 'ПБ' in plate_name and '56' in plate_name and '12' in plate_name and '8' in plate_name:
        plate_row = row
        print(f"✅ Найдена плита в строке {idx + 1}: {plate_name}")
        break

if plate_row is None:
    print("❌ Плита ПБ 56-12-8п не найдена в Excel")
    exit(1)

print()
print("📊 ДАННЫЕ ИЗ EXCEL (по колонкам):")
print()

# Показываем все колонки с данными
for col_idx in range(min(25, len(plate_row))):
    value = plate_row.iloc[col_idx]
    if pd.notna(value) and value != '':
        col_name = df.columns[col_idx] if col_idx < len(df.columns) else f"Col {col_idx}"
        print(f"   Колонка {col_idx} ({col_name}): {value}")

print()
print("=" * 80)
print("КЛЮЧЕВЫЕ КОЛОНКИ:")
print("=" * 80)
print()

# Ключевые колонки из документации
key_columns = {
    0: "Наименование",
    3: "Объем, м³",
    5: "Проволока, кг",
    6: "Канат, стоимость (руб)",
    13: "Бетон, стоимость (руб)",
    14: "Петли д 18",
    17: "Изоформ, кг",
    18: "Изоформ, стоимость (руб)",
    19: "Общая сумма (руб)"
}

for col_idx, description in key_columns.items():
    if col_idx < len(plate_row):
        value = plate_row.iloc[col_idx]
        if pd.notna(value):
            print(f"   Колонка {col_idx} - {description}: {value}")
        else:
            print(f"   Колонка {col_idx} - {description}: (пусто)")

print()
print("=" * 80)
print("АНАЛИЗ:")
print("=" * 80)
print()

# Проверяем колонку 6 (канат)
col6_value = plate_row.iloc[6] if 6 < len(plate_row) else None
print(f"Колонка 6 (канат): {col6_value}")
print()

# Проверяем колонку 19 (общая сумма)
col19_value = plate_row.iloc[19] if 19 < len(plate_row) else None
print(f"Колонка 19 (общая сумма): {col19_value}")
print()

# Проверяем, что может быть в желтой колонке
print("Возможные значения для желтой колонки (8104.5):")
print("  - Это может быть колонка с общей стоимостью армирования")
print("  - Или другая колонка (не колонка 6)")
print()

# Показываем все колонки с большими значениями (около 8000)
print("Колонки со значениями около 8000:")
for col_idx in range(len(plate_row)):
    value = plate_row.iloc[col_idx]
    if pd.notna(value) and isinstance(value, (int, float)):
        if 8000 <= value <= 9000:
            col_name = df.columns[col_idx] if col_idx < len(df.columns) else f"Col {col_idx}"
            print(f"   Колонка {col_idx} ({col_name}): {value}")

