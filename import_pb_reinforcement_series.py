#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Импорт данных "По серии" из Excel файла "Расчет новых цен на ПБ 10.09.2025 (1)"
в базу данных pb.db.

Структура данных:
- Для каждой плиты ПБ (17-99) берем значения "По серии" для нагрузок: 6, 8, 10, 12.5, 16, 21
- Марка бетона: до ПБ 59 включительно - М400, с ПБ 60 - M500
"""

import os
import re
import sqlite3
import sys
from pathlib import Path

# Настройка кодировки для Windows
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

try:
    import pandas as pd
except ImportError:
    print("[ОШИБКА] pandas не установлен. Установите: pip install pandas openpyxl")
    exit(1)

# Путь к базе данных
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DB = BASE_DIR / "pb.db"


def init_schema(db_path: Path | str = DEFAULT_DB) -> None:
    """Создает таблицу для хранения данных армирования 'По серии'"""
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Создаем таблицу для хранения данных "По серии" с маркой бетона
        cur.execute("""
            CREATE TABLE IF NOT EXISTS pb_reinforcement_series (
                length_dm INTEGER,
                load_code INTEGER,
                reinforcement_value REAL,
                concrete_grade TEXT,
                PRIMARY KEY(length_dm, load_code)
            )
        """)
        
        conn.commit()
        print("[OK] Таблица pb_reinforcement_series создана/проверена")
    finally:
        conn.close()


def extract_pb_number(beton_name: str) -> int | None:
    """Извлекает номер плиты ПБ из названия (например, 'ПБ 17' -> 17)"""
    if pd.isna(beton_name):
        return None
    
    beton_str = str(beton_name).strip()
    # Ищем паттерн "ПБ XX"
    match = re.search(r'ПБ\s*(\d+)', beton_str, re.IGNORECASE)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    return None


def get_concrete_grade(pb_number: int) -> str:
    """Определяет марку бетона по номеру плиты ПБ"""
    if pb_number <= 59:
        return "М400"
    else:
        return "M500"


def import_from_excel(xlsx_path: str | Path, db_path: Path | str = DEFAULT_DB) -> int:
    """
    Импортирует данные из Excel файла в базу данных.
    
    Лист "Свод армирования серия":
    - Строка 0-1: заголовки
    - Строка 2+: данные
    - Колонка 0: номер ПБ
    - Колонка 2: 6-нагрузка
    - Колонка 3: 8-нагрузка "По серии"
    - Колонка 5: 10-нагрузка "По серии"
    - Колонка 7: 12,5-нагрузка "По серии"
    - Колонка 9: 16-нагрузка "По серии"
    - Колонка 11: 21-нагрузка "По серии"
    """
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        print(f"[ОШИБКА] Файл не найден: {xlsx_path}")
        return 0
    
    print(f"[INFO] Читаю файл: {xlsx_path}")
    
    # Читаем Excel файл, лист "Свод армирования серия"
    try:
        df = pd.read_excel(xlsx_path, sheet_name='Свод армирования серия', header=None)
    except Exception as e:
        print(f"[ОШИБКА] Ошибка при чтении Excel: {e}")
        return 0
    
    print(f"[OK] Файл прочитан. Строк: {len(df)}")
    
    # Структура колонок (индексы):
    # Колонка 0: ПБ номер
    # Колонка 2: 6-нагрузка
    # Колонка 3: 8-нагрузка "По серии"
    # Колонка 5: 10-нагрузка "По серии"
    # Колонка 7: 12,5-нагрузка "По серии"
    # Колонка 9: 16-нагрузка "По серии"
    # Колонка 11: 21-нагрузка "По серии"
    
    load_columns = {
        6: 2,
        8: 3,
        10: 5,
        13: 7,  # 12.5 -> 13 в БД
        16: 9,
        21: 11
    }
    
    # Инициализируем схему БД
    init_schema(db_path)
    
    # Подготавливаем данные для вставки
    rows_to_insert = []
    
    # Начинаем со строки 2 (строки 0-1 — заголовки)
    for idx in range(2, len(df)):
        row = df.iloc[idx]
        
        # Извлекаем номер ПБ из колонки 0
        pb_number = extract_pb_number(row.get(0))
        
        if pb_number is None:
            continue
        
        # Определяем марку бетона
        concrete_grade = get_concrete_grade(pb_number)
        
        # Извлекаем значения для каждой нагрузки
        for load_code, col_idx in load_columns.items():
            value = row.get(col_idx)
            
            # Пропускаем пустые значения
            if pd.isna(value):
                continue
            
            try:
                reinforcement_value = float(str(value).replace(" ", "").replace(",", "."))
                rows_to_insert.append((pb_number, load_code, reinforcement_value, concrete_grade))
            except (ValueError, TypeError):
                continue
    
    if not rows_to_insert:
        print("[ОШИБКА] Не найдено данных для вставки")
        return 0
    
    print(f"[OK] Подготовлено {len(rows_to_insert)} записей для вставки")
    
    # Вставляем данные в БД
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Очищаем старые данные
        cur.execute("DELETE FROM pb_reinforcement_series")
        
        cur.executemany("""
            INSERT OR REPLACE INTO pb_reinforcement_series 
            (length_dm, load_code, reinforcement_value, concrete_grade)
            VALUES (?, ?, ?, ?)
        """, rows_to_insert)
        
        conn.commit()
        print(f"[OK] Успешно импортировано {len(rows_to_insert)} записей")
        
        # Показываем статистику
        cur.execute("SELECT COUNT(*) FROM pb_reinforcement_series")
        total_count = cur.fetchone()[0]
        print(f"[INFO] Всего записей в таблице: {total_count}")
        
        # Показываем примеры данных
        cur.execute("""
            SELECT length_dm, load_code, reinforcement_value, concrete_grade 
            FROM pb_reinforcement_series 
            ORDER BY length_dm, load_code 
            LIMIT 10
        """)
        print("\n[INFO] Примеры импортированных данных:")
        print("ПБ | Нагрузка | Армирование | Марка бетона")
        print("-" * 50)
        for row in cur.fetchall():
            print(f"ПБ {row[0]:2d} | {row[1]:2d} | {row[2]:8.1f} | {row[3]}")
        
        # Показываем границу М400/M500
        print("\n[INFO] Граница М400/M500:")
        cur.execute("""
            SELECT length_dm, load_code, reinforcement_value, concrete_grade 
            FROM pb_reinforcement_series 
            WHERE length_dm IN (58, 59, 60, 61)
            ORDER BY length_dm, load_code
        """)
        for row in cur.fetchall():
            print(f"ПБ {row[0]:2d} | {row[1]:2d} | {row[2]:8.1f} | {row[3]}")
        
    finally:
        conn.close()
    
    return len(rows_to_insert)


if __name__ == "__main__":
    
    # Путь к Excel файлу
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        # Ищем файл автоматически
        excel_path = None
        
        # Сначала ищем в корне проекта
        for file in BASE_DIR.glob("*.xlsx"):
            # Пропускаем временные файлы Excel (начинаются с ~$)
            if file.name.startswith("~$"):
                continue
            if "расчет" in file.name.lower() and "новых" in file.name.lower() and "цен" in file.name.lower():
                excel_path = file
                break
        
        # Если не нашли в корне, ищем в папке "банк знаний"
        if excel_path is None:
            bank_dir = BASE_DIR / "банк знаний"
            if bank_dir.exists():
                for file in bank_dir.glob("*.xlsx"):
                    # Пропускаем временные файлы Excel (начинаются с ~$)
                    if file.name.startswith("~$"):
                        continue
                    if "расчет" in file.name.lower() and "новых" in file.name.lower() and "цен" in file.name.lower():
                        excel_path = file
                        break
        
        if excel_path is None:
            print("[ОШИБКА] Файл Excel не найден автоматически.")
            print("[INFO] Использование: python import_pb_reinforcement_series.py <путь_к_excel_файлу>")
            print("[INFO] Или поместите файл 'Расчет новых цен на ПБ 10.09.2025 (1).xlsx' в корень проекта или в папку 'банк знаний'")
            sys.exit(1)
    
    print(f"[INFO] Начинаю импорт из файла: {excel_path}")
    count = import_from_excel(excel_path)
    
    if count > 0:
        print(f"\n[OK] Импорт завершен успешно! Импортировано записей: {count}")
    else:
        print("\n[ОШИБКА] Импорт не выполнен. Проверьте файл и структуру данных.")
