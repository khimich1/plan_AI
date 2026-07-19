#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Импорт данных о стоимости сырья и производственных расходов из Excel в БД
"""
import sqlite3
import pandas as pd
import os

def import_raw_material_costs(xlsx_path, db_path='pb.db'):
    """
    Импортирует данные о стоимости сырья и производственных расходов из Excel
    в таблицу raw_material_costs
    """
    # Сначала посмотрим какие листы доступны
    all_sheets = pd.read_excel(xlsx_path, sheet_name=None)
    print(f"Загружен Excel файл: {xlsx_path}")
    print(f"\nДоступные листы: {list(all_sheets.keys())}")
    
    # Ищем лист с прайсом (может называться по-разному)
    target_sheet = None
    for sheet_name in all_sheets.keys():
        if 'прайс' in sheet_name.lower() or 'price' in sheet_name.lower():
            target_sheet = sheet_name
            break
    
    if target_sheet is None:
        # Берем первый лист
        target_sheet = list(all_sheets.keys())[0]
    
    df = all_sheets[target_sheet]
    
    print(f"Используем лист: {target_sheet}")
    print(f"Всего строк: {len(df)}")
    print(f"Количество столбцов: {len(df.columns)}")
    
    # Ищем нужные столбцы
    name_col = None
    cost_col = None
    
    # Находим столбец с названиями
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if 'наименование' in col_lower or 'рсн' in col_lower:
            name_col = col
            break
    
    # Находим столбец "сырье+произ расходы"
    for col in df.columns:
        col_str = str(col).strip().lower()
        if 'сырье' in col_str and 'произ' in col_str and 'расход' in col_str:
            cost_col = col
            break
    
    if name_col is None or cost_col is None:
        print(f"\n[!] Ne naydeny nuzhnye stolbtsy!")
        print(f"name_col: {name_col}")
        print(f"cost_col: {cost_col}")
        return 0
    
    print(f"\n[OK] Nayden stolbets s nazvaniyami")
    print(f"[OK] Nayden stolbets so stoimostyu")
    
    # Собираем данные
    rows = []
    skipped = 0
    for idx, row in df.iterrows():
        name = str(row.get(name_col, '')).strip()
        cost = row.get(cost_col)
        
        # Проверяем, что есть название и стоимость
        if name and pd.notna(cost) and name.startswith('ПБ'):
            try:
                cost_value = float(str(cost).replace(' ', '').replace(',', '.'))
                rows.append((name, cost_value))
            except Exception as e:
                print(f"[!] Oshibka pri obrabotke stroki {idx}: {name}: {e}")
                skipped += 1
        else:
            if name and name.startswith('ПБ'):
                skipped += 1
    
    print(f"\nОбработано записей: {len(rows)}")
    print(f"Пропущено записей: {skipped}")
    
    # Создаем таблицу и заполняем данными
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Создаем таблицу
        cur.execute('''
            CREATE TABLE IF NOT EXISTS raw_material_costs (
                plate_name TEXT PRIMARY KEY,
                raw_material_and_production_cost REAL NOT NULL
            )
        ''')
        
        # Вставляем данные
        cur.executemany(
            'INSERT OR REPLACE INTO raw_material_costs (plate_name, raw_material_and_production_cost) VALUES (?, ?)',
            rows
        )
        
        conn.commit()
        print(f"\n[OK] Importirovano {len(rows)} zapisey v tablitsu raw_material_costs")
        
        # Показываем примеры
        cur.execute('SELECT * FROM raw_material_costs ORDER BY plate_name LIMIT 10')
        print("\nПервые 10 импортированных записей:")
        for row in cur.fetchall():
            print(f"  {row[0]}: {row[1]:.2f} руб.")
        
        # Показываем статистику
        cur.execute('SELECT COUNT(*), MIN(raw_material_and_production_cost), MAX(raw_material_and_production_cost), AVG(raw_material_and_production_cost) FROM raw_material_costs')
        stats = cur.fetchone()
        print(f"\nСтатистика:")
        print(f"  Всего записей: {stats[0]}")
        print(f"  Минимальная стоимость: {stats[1]:.2f} руб.")
        print(f"  Максимальная стоимость: {stats[2]:.2f} руб.")
        print(f"  Средняя стоимость: {stats[3]:.2f} руб.")
        
        return len(rows)
    finally:
        conn.close()


if __name__ == '__main__':
    xlsx_path = r'банк знаний\Расчет новых цен на ПБ 10.09.2025 (1).xlsx'
    
    if not os.path.exists(xlsx_path):
        print(f"[ERROR] Fayl ne nayden: {xlsx_path}")
    else:
        count = import_raw_material_costs(xlsx_path)
        print(f"\n{'='*60}")
        print(f"[GOTOVO] Importirovano {count} zapisey v bazu dannyh pb.db")
        print(f"{'='*60}")

