#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Детальный анализ логики расчета цен и себестоимости из файла
"Расчет новых цен на ПБ 10.09.2025 (1).xls"
"""

import os
import pandas as pd

file_path = os.path.join("банк знаний", "Расчет новых цен на ПБ 10.09.2025 (1).xls")

print("=" * 100)
print("АНАЛИЗ ЛОГИКИ РАСЧЕТА ЦЕН И СЕБЕСТОИМОСТИ ПЛИТ")
print("=" * 100)
print()

# Читаем основной лист с расчетами
try:
    df_main = pd.read_excel(file_path, sheet_name="Нов Серия для произв", engine='xlrd')
except:
    df_main = pd.read_excel(file_path, sheet_name="Нов Серия для произв", engine='openpyxl')

print("📊 ОСНОВНОЙ ЛИСТ: 'Нов Серия для произв'")
print("-" * 100)

# Показываем структуру колонок с понятными названиями
print("\n🔍 СТРУКТУРА РАСЧЕТА:")
print()

# Находим строку с заголовками (обычно строка 2-3)
header_row = None
for i in range(min(5, len(df_main))):
    row_vals = df_main.iloc[i].values
    if any('цемент' in str(v).lower() or 'песок' in str(v).lower() for v in row_vals if pd.notna(v)):
        header_row = i
        break

if header_row is not None:
    print(f"Заголовки найдены в строке {header_row + 1}:")
    for col_idx, col_name in enumerate(df_main.columns):
        val = df_main.iloc[header_row, col_idx]
        if pd.notna(val) and str(val).strip():
            print(f"  Колонка {col_idx}: {val}")

print("\n" + "=" * 100)
print("📋 КОМПОНЕНТЫ СЕБЕСТОИМОСТИ:")
print("=" * 100)

# Ищем строки с данными плит (начинаются с "ПБ")
plate_rows = df_main[df_main.iloc[:, 0].astype(str).str.startswith('ПБ', na=False)].copy()

if len(plate_rows) > 0:
    print(f"\nНайдено {len(plate_rows)} плит для анализа\n")
    
    # Берем первые 5 плит для детального анализа
    sample_plates = plate_rows.head(5)
    
    for idx, (_, row) in enumerate(sample_plates.iterrows(), 1):
        plate_name = row.iloc[0]
        length_dm = row.iloc[1] if pd.notna(row.iloc[1]) else "N/A"
        
        print(f"\n{'='*100}")
        print(f"ПЛИТА {idx}: {plate_name} (Длина: {length_dm} дм)")
        print(f"{'='*100}")
        
        # Показываем все значения в строке
        print("\nВсе компоненты расчета:")
        for col_idx, col_name in enumerate(df_main.columns):
            val = row.iloc[col_idx]
            if pd.notna(val):
                # Пропускаем пустые и NaN
                if str(val).strip() and str(val) != 'nan':
                    print(f"  [{col_idx}] {col_name}: {val}")
        
        print()

# Анализ колонок с ценами
print("\n" + "=" * 100)
print("💰 АНАЛИЗ ЦЕНОВЫХ КОЛОНОК:")
print("=" * 100)

price_cols = []
for col in df_main.columns:
    col_str = str(col).lower()
    if any(word in col_str for word in ['цена', 'прайс', 'стоимость', 'себестоимость', 'сумма', 'маржа']):
        price_cols.append(col)

print(f"\nНайдено {len(price_cols)} ценовых колонок:")
for col in price_cols:
    numeric_vals = pd.to_numeric(df_main[col], errors='coerce').dropna()
    if len(numeric_vals) > 0:
        print(f"\n  📌 {col}:")
        print(f"     Количество значений: {len(numeric_vals)}")
        print(f"     Диапазон: {numeric_vals.min():.2f} - {numeric_vals.max():.2f}")
        print(f"     Среднее: {numeric_vals.mean():.2f}")
        
        # Показываем примеры для первых плит
        sample = numeric_vals.head(3).tolist()
        print(f"     Примеры: {sample}")

# Анализ материалов
print("\n" + "=" * 100)
print("🧱 АНАЛИЗ МАТЕРИАЛОВ:")
print("=" * 100)

material_keywords = ['цемент', 'песок', 'щебень', 'бетон', 'армирование', 'проволока', 'канат', 'петли', 'изоформ']
for keyword in material_keywords:
    matching_cols = [col for col in df_main.columns if keyword in str(col).lower()]
    if matching_cols:
        print(f"\n  🔹 {keyword.upper()}:")
        for col in matching_cols:
            numeric_vals = pd.to_numeric(df_main[col], errors='coerce').dropna()
            if len(numeric_vals) > 0:
                print(f"     {col}: {len(numeric_vals)} значений, примеры: {numeric_vals.head(3).tolist()}")

# Анализ коэффициентов и маржи
print("\n" + "=" * 100)
print("📈 АНАЛИЗ КОЭФФИЦИЕНТОВ И МАРЖИ:")
print("=" * 100)

coef_cols = [col for col in df_main.columns if any(word in str(col).lower() for word in ['коэф', 'маржа', 'скидк', '%'])]
if coef_cols:
    print(f"\nНайдено {len(coef_cols)} колонок с коэффициентами/маржой:")
    for col in coef_cols:
        numeric_vals = pd.to_numeric(df_main[col], errors='coerce').dropna()
        if len(numeric_vals) > 0:
            print(f"  {col}:")
            print(f"    Диапазон: {numeric_vals.min():.4f} - {numeric_vals.max():.4f}")
            print(f"    Среднее: {numeric_vals.mean():.4f}")

# Попытка найти формулы расчета
print("\n" + "=" * 100)
print("🔬 ПОПЫТКА ВЫЯВИТЬ ФОРМУЛЫ РАСЧЕТА:")
print("=" * 100)

# Ищем колонку "Общая сумма" и пытаемся понять, из чего она складывается
if 'Общая сумма' in df_main.columns:
    print("\nКолонка 'Общая сумма' найдена!")
    # Показываем примеры расчетов
    for idx, (_, row) in enumerate(plate_rows.head(3).iterrows(), 1):
        plate_name = row.iloc[0]
        total_sum = row['Общая сумма'] if 'Общая сумма' in row.index and pd.notna(row['Общая сумма']) else None
        
        if total_sum:
            print(f"\n  {plate_name}:")
            print(f"    Общая сумма: {total_sum:.2f}")
            
            # Пытаемся найти компоненты, которые могли войти в сумму
            # Обычно это материалы до колонки "Общая сумма"
            total_col_idx = list(df_main.columns).index('Общая сумма')
            print(f"    Возможные компоненты (до колонки 'Общая сумма'):")
            for i in range(min(5, total_col_idx)):
                val = row.iloc[i]
                if pd.notna(val) and isinstance(val, (int, float)):
                    print(f"      [{i}] {df_main.columns[i]}: {val}")

print("\n" + "=" * 100)
print("✅ АНАЛИЗ ЗАВЕРШЕН")
print("=" * 100)

