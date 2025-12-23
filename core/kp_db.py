#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
База данных для хранения коммерческих предложений (КП) в формате XLSX

База данных: plita.db

Структура таблиц:
- KP_offers: основная информация о КП
- kp_plates: позиции (плиты) в каждом КП
- kp_files: файлы XLSX (хранит сам файл как BLOB и путь)
- kp_meta: метаданные (статус КП)
"""

import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional, Tuple

# Путь к базе данных (в корне проекта)
DEFAULT_DB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'plita.db')


def init_schema(db_path: str = DEFAULT_DB) -> None:
    """
    Создаёт таблицы в базе данных, если их ещё нет.
    
    Простыми словами:
    - Проверяет, есть ли таблицы в БД
    - Если нет — создаёт их с нужными колонками
    - Это как создать пустую таблицу Excel с заголовками
    """
    conn = sqlite3.connect(db_path)
    try:
        # КРИТИЧНО: Включаем поддержку FOREIGN KEY (по умолчанию выключена в SQLite)
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Таблица 1: KP_offers - Основная информация о КП
        # PRIMARY KEY - порядковый номер КП (auto increment)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS KP_offers (
                kp_id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_date TEXT NOT NULL,
                customer_name TEXT,
                manager_name TEXT,
                discount_percent REAL DEFAULT 0,
                subtotal REAL,
                vat_amount REAL,
                total_amount REAL,
                delivery_conditions TEXT,
                payment_conditions TEXT,
                execution_terms TEXT
            )
        ''')
        
        # Таблица 2: kp_plates - Позиции (плиты) в каждом КП
        # kp_id - связь с KP_offers
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_plates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER,
                plate_name TEXT NOT NULL,
                length_m REAL,
                width_m REAL,
                load_class INTEGER,
                qty INTEGER NOT NULL,
                unit_weight REAL,
                total_weight REAL,
                discounted_price REAL,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        
        # Таблица 3: kp_files - Файлы XLSX
        # kp_id - связь с KP_offers
        # xlsx_file - сам файл как BLOB (двоичные данные)
        # file_path - путь к файлу
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                xlsx_file BLOB,
                file_path TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE,
                UNIQUE(kp_id)
            )
        ''')
        
        # Таблица 4: kp_meta - Метаданные
        # kp_id - связь с KP_offers
        # status - статус КП: выполнено, отклонено, в работе, в ожидании
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_meta (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                status TEXT DEFAULT 'в работе',
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE,
                UNIQUE(kp_id)
            )
        ''')
        
        # Создаём индексы для быстрого поиска
        # Это как закладки в книге — помогают быстро найти нужную информацию
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_plates ON kp_plates(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_files ON kp_files(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_meta ON kp_meta(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_meta_status ON kp_meta(status)')
        
        conn.commit()
    finally:
        conn.close()


def save_kp_to_db(
    creation_date: str,
    order_data: List[Dict],
    xlsx_file_path: Optional[str] = None,
    customer_name: Optional[str] = None,
    manager_name: Optional[str] = None,
    discount_percent: float = 0,
    delivery_conditions: Optional[str] = None,
    payment_conditions: Optional[str] = None,
    execution_terms: Optional[str] = None,
    status: str = 'в работе',
    db_path: str = DEFAULT_DB
) -> int:
    """
    Сохраняет КП в базу данных.
    
    Простыми словами:
    - Берёт всю информацию о КП
    - Записывает её в базу данных
    - Возвращает порядковый номер КП (kp_id)
    
    Аргументы:
        creation_date: дата создания КП (например: "01.01.2024" или "2024-01-01")
        order_data: список плит в заказе (из функции generate_commercial_offer_xlsx)
        xlsx_file_path: путь к файлу XLSX (если есть)
        customer_name: имя клиента
        manager_name: имя менеджера
        discount_percent: процент скидки (0-100)
        delivery_conditions: условия поставки
        payment_conditions: условия оплаты
        execution_terms: сроки выполнения
        status: статус КП (по умолчанию "в работе")
        db_path: путь к базе данных
    
    Возвращает:
        Порядковый номер КП (kp_id)
    """
    init_schema(db_path)
    
    # 🔥 ИСПРАВЛЕНИЕ: Используем ту же функцию расчета, что и в XLSX
    # Это гарантирует, что суммы в БД и в XLSX файле будут одинаковыми
    try:
        from core.commercial_offer_xlsx import calculate_total_cost
        totals = calculate_total_cost(order_data, discount_percent)
        subtotal = totals['subtotal']
        vat_amount = totals['vat_amount']
        total_amount = totals['total_with_vat']
    except ImportError:
        # Fallback: старая логика, если модуль не найден
        subtotal = 0.0
        for item in order_data:
            qty = item.get('qty', 0)
            unit_price = item.get('unit_price', 0.0)
            discounted_price = unit_price * (1 - discount_percent / 100)
            subtotal += discounted_price * qty
        
        vat_amount = round(subtotal * 0.20, 2)
        total_amount = round(subtotal + vat_amount, 2)
    
    conn = sqlite3.connect(db_path)
    try:
        # Включаем поддержку FOREIGN KEY
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Сохраняем основную информацию о КП
        cur.execute('''
            INSERT INTO KP_offers (
                creation_date, customer_name, manager_name, discount_percent,
                subtotal, vat_amount, total_amount,
                delivery_conditions, payment_conditions, execution_terms
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            creation_date, customer_name, manager_name, discount_percent,
            subtotal, vat_amount, total_amount,
            delivery_conditions, payment_conditions, execution_terms
        ))
        
        # Получаем порядковый номер созданного КП
        kp_id = cur.lastrowid
        
        # Сохраняем позиции (плиты)
        for idx, item in enumerate(order_data, start=1):
            qty = item.get('qty', 0)
            unit_price = item.get('unit_price', 0.0)
            discounted_price = unit_price * (1 - discount_percent / 100)
            weight = item.get('weight', 0.0)
            unit_weight = weight / qty if qty > 0 else 0.0
            
            cur.execute('''
                INSERT INTO kp_plates (
                    kp_id, position_number, plate_name,
                    length_m, width_m, load_class,
                    qty, unit_weight, total_weight, discounted_price
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                kp_id, idx, item.get('name', ''),
                item.get('length_m', 0), item.get('width_m', 0), item.get('load_class', 800),
                qty, unit_weight, weight, discounted_price
            ))
        
        # Сохраняем файл XLSX (если указан путь)
        if xlsx_file_path and os.path.exists(xlsx_file_path):
            with open(xlsx_file_path, 'rb') as f:
                xlsx_blob = f.read()
            
            cur.execute('''
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
            ''', (kp_id, xlsx_blob, xlsx_file_path))
        elif xlsx_file_path:
            # Если путь указан, но файла нет, сохраняем только путь
            cur.execute('''
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
            ''', (kp_id, None, xlsx_file_path))
        
        # Сохраняем метаданные (статус)
        cur.execute('''
            INSERT INTO kp_meta (kp_id, status)
            VALUES (?, ?)
        ''', (kp_id, status))
        
        conn.commit()
        return kp_id
    
    finally:
        conn.close()


def get_kp_by_id(kp_id: int, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """
    Получает информацию о КП по порядковому номеру.
    
    Простыми словами:
    - Ищет КП по номеру в базе
    - Возвращает всю информацию о нём + список плит + файл + статус
    
    Возвращает:
        Словарь с информацией о КП или None, если не найдено
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Получаем основную информацию
        cur.execute('SELECT * FROM KP_offers WHERE kp_id = ?', (kp_id,))
        row = cur.fetchone()
        if not row:
            return None
        
        kp_data = dict(row)
        
        # Получаем список плит
        cur.execute('''
            SELECT * FROM kp_plates 
            WHERE kp_id = ? 
            ORDER BY position_number
        ''', (kp_id,))
        plates = [dict(row) for row in cur.fetchall()]
        kp_data['plates'] = plates
        
        # Получаем информацию о файле
        cur.execute('SELECT * FROM kp_files WHERE kp_id = ?', (kp_id,))
        file_row = cur.fetchone()
        if file_row:
            file_data = dict(file_row)
            # BLOB не нужно конвертировать в строку, оставляем как bytes
            kp_data['file'] = file_data
        
        # Получаем статус из метаданных
        cur.execute('SELECT status FROM kp_meta WHERE kp_id = ?', (kp_id,))
        meta_row = cur.fetchone()
        if meta_row:
            kp_data['status'] = meta_row['status']
        
        return kp_data
    
    finally:
        conn.close()


def get_all_kp_by_status(status: str, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все КП с определённым статусом.
    
    Простыми словами:
    - Ищет все КП со статусом (например, "в работе")
    - Возвращает список таких КП
    
    Аргументы:
        status: статус для поиска ("выполнено", "отклонено", "в работе", "в ожидании")
    
    Возвращает:
        Список словарей с информацией о КП
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Получаем ID всех КП с нужным статусом
        cur.execute('SELECT kp_id FROM kp_meta WHERE status = ?', (status,))
        kp_ids = [row['kp_id'] for row in cur.fetchall()]
        
        # Получаем информацию о каждом КП
        result = []
        for kp_id in kp_ids:
            kp_data = get_kp_by_id(kp_id, db_path)
            if kp_data:
                result.append(kp_data)
        
        return result
    
    finally:
        conn.close()


def update_kp_status(kp_id: int, new_status: str, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет статус КП.
    
    Простыми словами:
    - Меняет статус КП (например, с "в работе" на "выполнено")
    
    Аргументы:
        kp_id: порядковый номер КП
        new_status: новый статус ("выполнено", "отклонено", "в работе", "в ожидании")
    
    Возвращает:
        True если успешно, False если КП не найдено
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        cur.execute('''
            UPDATE kp_meta 
            SET status = ?
            WHERE kp_id = ?
        ''', (new_status, kp_id))
        
        conn.commit()
        return cur.rowcount > 0
    
    finally:
        conn.close()


def save_xlsx_file(kp_id: int, xlsx_file_path: str, db_path: str = DEFAULT_DB) -> bool:
    """
    Сохраняет файл XLSX в базу данных (обновляет или создаёт запись).
    
    Простыми словами:
    - Читает файл XLSX с диска
    - Сохраняет его в базу данных как двоичные данные
    - Также сохраняет путь к файлу
    
    Возвращает:
        True если успешно, False если ошибка
    """
    init_schema(db_path)
    
    if not os.path.exists(xlsx_file_path):
        return False
    
    conn = sqlite3.connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        with open(xlsx_file_path, 'rb') as f:
            xlsx_blob = f.read()
        
        cur = conn.cursor()
        # Проверяем, есть ли уже запись
        cur.execute('SELECT id FROM kp_files WHERE kp_id = ?', (kp_id,))
        exists = cur.fetchone()
        
        if exists:
            # Обновляем существующую запись
            cur.execute('''
                UPDATE kp_files 
                SET xlsx_file = ?, file_path = ?
                WHERE kp_id = ?
            ''', (xlsx_blob, xlsx_file_path, kp_id))
        else:
            # Создаём новую запись
            cur.execute('''
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
            ''', (kp_id, xlsx_blob, xlsx_file_path))
        
        conn.commit()
        return True
    
    finally:
        conn.close()


def get_xlsx_file(kp_id: int, output_path: Optional[str] = None, db_path: str = DEFAULT_DB) -> Optional[bytes]:
    """
    Получает файл XLSX из базы данных.
    
    Простыми словами:
    - Извлекает файл XLSX из базы данных
    - Если указан output_path, сохраняет файл на диск
    - Иначе возвращает файл как bytes (двоичные данные)
    
    Аргументы:
        kp_id: порядковый номер КП
        output_path: путь для сохранения файла (опционально)
        db_path: путь к базе данных
    
    Возвращает:
        bytes (двоичные данные файла) или None если не найдено
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        cur.execute('SELECT xlsx_file FROM kp_files WHERE kp_id = ?', (kp_id,))
        row = cur.fetchone()
        
        if not row or not row[0]:
            return None
        
        xlsx_data = row[0]
        
        # Если указан путь для сохранения, сохраняем файл
        if output_path:
            with open(output_path, 'wb') as f:
                f.write(xlsx_data)
        
        return xlsx_data
    
    finally:
        conn.close()


def delete_kp_by_id(kp_id: int, db_path: str = DEFAULT_DB) -> bool:
    """
    Удаляет КП из базы данных по порядковому номеру.
    
    Простыми словами:
    - Ищет КП по номеру
    - Удаляет его из базы (включая все плиты, файлы, метаданные)
    - Благодаря CASCADE DELETE удаляются все связанные записи автоматически
    
    Аргументы:
        kp_id: порядковый номер КП для удаления
        db_path: путь к базе данных
    
    Возвращает:
        True если КП был найден и удалён, False если КП не найдено
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        # КРИТИЧНО: Включаем поддержку FOREIGN KEY (по умолчанию выключена в SQLite)
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Проверяем, существует ли КП
        cur.execute('SELECT kp_id FROM KP_offers WHERE kp_id = ?', (kp_id,))
        if not cur.fetchone():
            return False
        
        # Удаляем КП (CASCADE автоматически удалит все связанные записи)
        cur.execute('DELETE FROM KP_offers WHERE kp_id = ?', (kp_id,))
        
        conn.commit()
        print(f"[DB] ✅ КП #{kp_id} успешно удалено из базы данных")
        return True
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при удалении КП #{kp_id}: {e}")
        return False
    
    finally:
        conn.close()


def clear_all_kp(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """
    Полностью очищает все таблицы с КП из базы данных.
    
    Простыми словами:
    - Удаляет ВСЕ КП из базы данных
    - Очищает все таблицы: KP_offers, kp_plates, kp_files, kp_meta
    - Сбрасывает счётчики AUTOINCREMENT, чтобы новые КП начинались с 1
    - Это как стереть все записи из всех таблиц Excel и начать заново
    
    Возвращает:
        Словарь с количеством удалённых записей из каждой таблицы
    """
    init_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        # Включаем поддержку FOREIGN KEY
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Подсчитываем количество записей перед удалением
        cur.execute('SELECT COUNT(*) FROM KP_offers')
        kp_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_plates')
        plates_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_files')
        files_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta')
        meta_count = cur.fetchone()[0]
        
        # Удаляем все записи из всех таблиц
        # Порядок важен: сначала зависимые таблицы, потом основную
        cur.execute('DELETE FROM kp_plates')
        cur.execute('DELETE FROM kp_files')
        cur.execute('DELETE FROM kp_meta')
        cur.execute('DELETE FROM KP_offers')
        
        # Сбрасываем счётчики AUTOINCREMENT
        # Это нужно, чтобы новые КП начинались с номера 1
        cur.execute('DELETE FROM sqlite_sequence WHERE name IN (?, ?, ?, ?)', 
                   ('KP_offers', 'kp_plates', 'kp_files', 'kp_meta'))
        
        conn.commit()
        
        result = {
            'kp_offers': kp_count,
            'kp_plates': plates_count,
            'kp_files': files_count,
            'kp_meta': meta_count
        }
        
        print(f"[DB] ✅ Полная очистка БД завершена:")
        print(f"  - Удалено КП: {kp_count}")
        print(f"  - Удалено записей плит: {plates_count}")
        print(f"  - Удалено файлов: {files_count}")
        print(f"  - Удалено метаданных: {meta_count}")
        
        return result
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при полной очистке БД: {e}")
        import traceback
        traceback.print_exc()
        raise
    
    finally:
        conn.close()


# ==================== ПРИМЕР ИСПОЛЬЗОВАНИЯ ====================

if __name__ == '__main__':
    # Пример тестовых данных
    test_order_data = [
        {
            'name': 'ПБ 78-12-8п',
            'length_m': 7.8,
            'width_m': 1.2,
            'qty': 4,
            'load_class': 800,
            'unit_price': 5000.0,
            'weight': 987.84
        },
        {
            'name': 'ПБ 66-3-8п',
            'length_m': 6.6,
            'width_m': 0.3,
            'qty': 2,
            'load_class': 800,
            'unit_price': 1200.0,
            'weight': 104.54
        }
    ]
    
    print("📝 Создаю тестовую БД...")
    
    # Пример сохранения КП в БД
    kp_id = save_kp_to_db(
        creation_date='01.01.2024',
        order_data=test_order_data,
        xlsx_file_path=None,  # Можно указать путь к файлу
        customer_name='ООО Тест',
        manager_name='Иванов И.И.',
        discount_percent=5.0,
        delivery_conditions='Самовывоз',
        payment_conditions='Предоплата 100%',
        execution_terms='14 дней',
        status='в работе'
    )
    print(f"✅ КП сохранён в БД с порядковым номером: {kp_id}")
    
    # Пример получения КП
    print("\n🔍 Ищу КП по номеру...")
    kp = get_kp_by_id(kp_id)
    if kp:
        print(f"✅ Найден КП № {kp['kp_id']}")
        print(f"   Дата: {kp['creation_date']}")
        print(f"   Клиент: {kp['customer_name']}")
        print(f"   Менеджер: {kp['manager_name']}")
        print(f"   Сумма: {kp['total_amount']:.2f} ₽")
        print(f"   Статус: {kp.get('status', 'не указан')}")
        print(f"   Плит в заказе: {len(kp['plates'])} позиций")
    
    # Пример получения всех КП в работе
    print("\n📋 Все КП в работе:")
    all_kp = get_all_kp_by_status('в работе')
    for kp in all_kp:
        print(f"  • КП № {kp['kp_id']} - {kp['customer_name']} - {kp['total_amount']:.2f} ₽")
