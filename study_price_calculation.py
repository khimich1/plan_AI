#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Изучение файла "Расчет новых цен на ПБ 10.09.2025 (1).xls"
Анализ структуры и логики расчета цен и себестоимости
"""

import os
import pandas as pd

# Путь к файлу
file_path = os.path.join("банк знаний", "Расчет новых цен на ПБ 10.09.2025 (1).xls")

if not os.path.exists(file_path):
    print(f"❌ Файл не найден: {file_path}")
    exit(1)

print(f"📄 Изучаем файл: {file_path}\n")

# Читаем все листы
try:
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='xlrd')
    print(f"✅ Найдено листов: {len(all_sheets)}")
    print(f"📋 Названия листов: {list(all_sheets.keys())}\n")
except Exception as e:
    print(f"⚠️ Ошибка чтения с xlrd: {e}")
    # Попробуем с другим движком
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        print(f"✅ Найдено листов: {len(all_sheets)}")
        print(f"📋 Названия листов: {list(all_sheets.keys())}\n")
    except Exception as e2:
        print(f"❌ Ошибка с openpyxl: {e2}")
        exit(1)

# Изучаем каждый лист
for sheet_name, df in all_sheets.items():
    print("=" * 80)
    print(f"📊 ЛИСТ: {sheet_name}")
    print("=" * 80)
    print(f"Размер: {df.shape[0]} строк × {df.shape[1]} колонок\n")
    
    # Показываем колонки
    print("Колонки:")
    for i, col in enumerate(df.columns):
        print(f"  {i}: {repr(col)}")
    print()
    
    # Показываем первые строки
    print("Первые 20 строк:")
    print(df.head(20).to_string())
    print()
    
    # Ищем формулы или расчетные поля
    print("Поиск числовых колонок (возможно, содержат расчеты):")
    numeric_cols = df.select_dtypes(include=['number']).columns
    for col in numeric_cols:
        non_null = df[col].notna().sum()
        if non_null > 0:
            print(f"  {col}: {non_null} непустых значений")
            sample_vals = df[col].dropna().head(5).tolist()
            print(f"    Примеры: {sample_vals}")
    print()
    
    # Ищем колонки с названиями плит
    print("Поиск колонок с наименованиями плит:")
    for col in df.columns:
        col_str = str(col).lower()
        if any(word in col_str for word in ['наимен', 'назв', 'плит', 'пб', 'name']):
            print(f"  Найдена колонка: {col}")
            # Показываем примеры значений
            non_empty = df[col].dropna().head(10)
            if len(non_empty) > 0:
                print(f"    Примеры значений:")
                for val in non_empty:
                    print(f"      - {val}")
    print()
    
    # Ищем колонки с ценами/стоимостью
    print("Поиск колонок с ценами/стоимостью:")
    for col in df.columns:
        col_str = str(col).lower()
        if any(word in col_str for word in ['цена', 'стоимость', 'себестоимость', 'price', 'cost', 'руб']):
            print(f"  Найдена колонка: {col}")
            # Показываем статистику
            numeric_vals = pd.to_numeric(df[col], errors='coerce').dropna()
            if len(numeric_vals) > 0:
                print(f"    Непустых значений: {len(numeric_vals)}")
                print(f"    Минимум: {numeric_vals.min():.2f}")
                print(f"    Максимум: {numeric_vals.max():.2f}")
                print(f"    Среднее: {numeric_vals.mean():.2f}")
    print()
    
    # Ищем колонки с нагрузками
    print("Поиск колонок с нагрузками:")
    for col in df.columns:
        col_str = str(col).lower()
        if any(word in col_str for word in ['нагруз', 'load', '6', '8', '10', '12']):
            print(f"  Найдена колонка: {col}")
    print()
    
    # Ищем колонки с размерами (длина, ширина)
    print("Поиск колонок с размерами (длина, ширина):")
    for col in df.columns:
        col_str = str(col).lower()
        if any(word in col_str for word in ['длина', 'ширина', 'length', 'width', 'размер']):
            print(f"  Найдена колонка: {col}")
            non_empty = df[col].dropna().head(5)
            if len(non_empty) > 0:
                print(f"    Примеры: {non_empty.tolist()}")
    print()
    
    print("\n" + "=" * 80 + "\n")

