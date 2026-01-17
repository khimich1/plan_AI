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
- completed_plates: выполненные плиты (перенесены из kp_plates после завершения дня)
- plate_rests: остатки от резки плит
- managers: менеджеры (ФИО, контактный номер, email)
"""

import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional, Tuple

# Путь к базе данных (в корне проекта)
DEFAULT_DB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'plita.db')


def _connect(db_path: str) -> sqlite3.Connection:
    """
    Безопасное подключение к SQLite (plita.db).

    - WAL уменьшает риск повреждения базы при сбоях/перезапусках.
    - foreign_keys нужен для корректной работы связей между таблицами.
    """
    conn = sqlite3.connect(db_path)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


def init_schema(db_path: str = DEFAULT_DB) -> None:
    """
    Создаёт таблицы в базе данных, если их ещё нет.
    
    Простыми словами:
    - Проверяет, есть ли таблицы в БД
    - Если нет — создаёт их с нужными колонками
    - Это как создать пустую таблицу Excel с заголовками
    """
    conn = _connect(db_path)
    try:
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
        
        # Таблица 5: completed_plates - Выполненные плиты
        # Сюда переносятся плиты после завершения дня производства
        cur.execute('''
            CREATE TABLE IF NOT EXISTS completed_plates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                plate_name TEXT NOT NULL,
                length_m REAL,
                width_m REAL,
                load_class INTEGER,
                qty INTEGER NOT NULL,
                completed_date TEXT NOT NULL,
                production_day INTEGER,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        
        # Таблица 6: plate_rests - Остатки от резки плит
        # Хранит информацию об остатках, которые образуются при продольном резе
        # Статусы: available (доступен), used (использован), completed (выполнен), discarded (списан)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS plate_rests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                source_plate_name TEXT NOT NULL,
                rest_width_mm INTEGER NOT NULL,
                length_m REAL NOT NULL,
                qty INTEGER NOT NULL DEFAULT 1,
                status TEXT DEFAULT 'available',
                created_date TEXT NOT NULL,
                used_date TEXT,
                production_day INTEGER,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        
        # Таблица 7: managers - Менеджеры
        # Хранит информацию о менеджерах: ФИО, контактный номер, email
        cur.execute('''
            CREATE TABLE IF NOT EXISTS managers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fio TEXT NOT NULL,
                contact_number TEXT NOT NULL,
                email TEXT NOT NULL,
                UNIQUE(email)
            )
        ''')
        
        # Создаём индексы для быстрого поиска
        # Это как закладки в книге — помогают быстро найти нужную информацию
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_plates ON kp_plates(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_files ON kp_files(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_meta ON kp_meta(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_meta_status ON kp_meta(status)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_completed_kp_id ON completed_plates(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_completed_date ON completed_plates(completed_date)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_rests_kp_id ON plate_rests(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_rests_status ON plate_rests(status)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_managers_email ON managers(email)')
        
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
        
        vat_amount = round(subtotal * 0.22, 2)
        total_amount = round(subtotal + vat_amount, 2)
    
    conn = _connect(db_path)
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
    conn = _connect(db_path)
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
    conn = _connect(db_path)
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
    conn = _connect(db_path)
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
    
    conn = _connect(db_path)
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
    conn = _connect(db_path)
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
    conn = _connect(db_path)
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


def clear_all_plates_data(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """
    Полностью очищает ВСЕ данные о плитах из базы данных.
    
    Простыми словами:
    - Удаляет ВСЕ КП (и в работе, и выполненные, и отклонённые)
    - Удаляет ВСЕ плиты (и невыполненные, и выполненные)
    - Удаляет ВСЕ остатки от резки
    - Сбрасывает счётчики AUTOINCREMENT
    - Это как полностью очистить завод от всех заказов и начать с нуля
    
    ⚠️ ВНИМАНИЕ: Это необратимая операция!
    
    Возвращает:
        Словарь с количеством удалённых записей из каждой таблицы
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        # Включаем поддержку FOREIGN KEY
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Подсчитываем количество записей перед удалением
        cur.execute('SELECT COUNT(*) FROM KP_offers')
        kp_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_plates')
        plates_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM completed_plates')
        completed_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM plate_rests')
        rests_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_files')
        files_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta')
        meta_count = cur.fetchone()[0]
        
        # Удаляем все записи из всех таблиц
        # Порядок важен: сначала зависимые таблицы, потом основную
        cur.execute('DELETE FROM kp_plates')
        cur.execute('DELETE FROM completed_plates')
        cur.execute('DELETE FROM plate_rests')
        cur.execute('DELETE FROM kp_files')
        cur.execute('DELETE FROM kp_meta')
        cur.execute('DELETE FROM KP_offers')
        
        # Сбрасываем счётчики AUTOINCREMENT
        cur.execute('''
            DELETE FROM sqlite_sequence 
            WHERE name IN ('KP_offers', 'kp_plates', 'completed_plates', 
                          'plate_rests', 'kp_files', 'kp_meta')
        ''')
        
        conn.commit()
        
        result = {
            'kp_offers': kp_count,
            'kp_plates': plates_count,
            'completed_plates': completed_count,
            'plate_rests': rests_count,
            'kp_files': files_count,
            'kp_meta': meta_count,
            'total': kp_count + plates_count + completed_count + rests_count + files_count + meta_count
        }
        
        print(f"[DB] ✅ ПОЛНАЯ ОЧИСТКА ВСЕХ ДАННЫХ О ПЛИТАХ:")
        print(f"  - Удалено КП: {kp_count}")
        print(f"  - Удалено плит в работе: {plates_count}")
        print(f"  - Удалено выполненных плит: {completed_count}")
        print(f"  - Удалено остатков: {rests_count}")
        print(f"  - Удалено файлов: {files_count}")
        print(f"  - Удалено метаданных: {meta_count}")
        print(f"  - ВСЕГО ЗАПИСЕЙ: {result['total']}")
        
        return result
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при полной очистке БД: {e}")
        import traceback
        traceback.print_exc()
        conn.rollback()
        raise
    
    finally:
        conn.close()


def get_db_stats(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """
    Получает статистику по базе данных.
    
    Возвращает:
        Словарь с количеством записей в каждой таблице
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        
        cur.execute('SELECT COUNT(*) FROM KP_offers')
        kp_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_plates')
        plates_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM completed_plates')
        completed_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM plate_rests')
        rests_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta WHERE status = "в работе"')
        in_work_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta WHERE status = "выполнено"')
        completed_kp_count = cur.fetchone()[0]
        
        return {
            'kp_total': kp_count,
            'kp_in_work': in_work_count,
            'kp_completed': completed_kp_count,
            'plates_in_work': plates_count,
            'plates_completed': completed_count,
            'plate_rests': rests_count
        }
    
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
    conn = _connect(db_path)
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


# ==================== ФУНКЦИИ ДЛЯ ВЫПОЛНЕННЫХ ПЛИТ ====================

def move_plates_to_completed(
    kp_id: int,
    plates_to_complete: List[Dict],
    production_day: int,
    db_path: str = DEFAULT_DB
) -> int:
    """
    Переносит плиты из kp_plates в completed_plates.
    
    Простыми словами:
    - Берёт список плит, которые нужно отметить как выполненные
    - Записывает их в таблицу completed_plates
    - Уменьшает количество в kp_plates (или удаляет, если qty = 0)
    
    Аргументы:
        kp_id: номер КП, к которому относятся плиты
        plates_to_complete: список плит для переноса (словари с plate_name, qty и т.д.)
        production_day: номер дня производства
        db_path: путь к базе данных
    
    Возвращает:
        Количество перенесённых плит (сумма qty)
    """
    init_schema(db_path)
    conn = _connect(db_path)
    completed_count = 0
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        completed_date = datetime.now().strftime('%d.%m.%Y')
        
        for plate in plates_to_complete:
            plate_name = plate.get('plate_name', '')
            qty = plate.get('qty', 1)
            
            if not plate_name:
                continue
            
            # Вставляем в completed_plates
            cur.execute('''
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class,
                    qty, completed_date, production_day
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                kp_id,
                plate_name,
                plate.get('length_m', 0),
                plate.get('width_m', 0),
                plate.get('load_class', 800),
                qty,
                completed_date,
                production_day
            ))
            
            # Уменьшаем qty в kp_plates
            cur.execute('''
                UPDATE kp_plates 
                SET qty = qty - ?
                WHERE kp_id = ? AND plate_name = ? AND qty > 0
            ''', (qty, kp_id, plate_name))
            
            completed_count += qty
        
        # Удаляем записи с qty <= 0
        cur.execute('DELETE FROM kp_plates WHERE qty <= 0')
        
        conn.commit()
        print(f"[DB] ✅ Перенесено {completed_count} плит в completed_plates (КП #{kp_id}, день {production_day})")
        return completed_count
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при переносе плит: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


# ==================== ФУНКЦИИ ДЛЯ ОСТАТКОВ ПЛИТ ====================

def create_plate_rest(
    kp_id: int,
    source_plate_name: str,
    rest_width_mm: int,
    length_m: float,
    production_day: int,
    qty: int = 1,
    db_path: str = DEFAULT_DB
) -> int:
    """
    Создает запись об остатке плиты.
    
    Простыми словами:
    - При продольном резе плиты образуется остаток
    - Эта функция сохраняет информацию об остатке в БД
    - Остаток можно использовать для других заказов
    
    Аргументы:
        kp_id: номер КП, при выполнении которого образовался остаток
        source_plate_name: имя исходной плиты (из которой вырезали)
        rest_width_mm: ширина остатка в мм
        length_m: длина остатка в метрах
        production_day: номер дня производства
        qty: количество остатков (по умолчанию 1)
        db_path: путь к базе данных
    
    Возвращает:
        ID созданной записи или 0 при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        created_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            INSERT INTO plate_rests (
                kp_id, source_plate_name, rest_width_mm, length_m,
                qty, status, created_date, production_day
            ) VALUES (?, ?, ?, ?, ?, 'available', ?, ?)
        ''', (
            kp_id,
            source_plate_name,
            rest_width_mm,
            length_m,
            qty,
            created_date,
            production_day
        ))
        
        rest_id = cur.lastrowid
        conn.commit()
        print(f"[DB] ✅ Создан остаток #{rest_id}: {rest_width_mm}мм x {length_m}м (КП #{kp_id})")
        return rest_id
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при создании остатка: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def get_available_rests(
    kp_id: int = None,
    db_path: str = DEFAULT_DB
) -> List[Dict]:
    """
    Возвращает список доступных остатков.
    
    Простыми словами:
    - Получает все остатки со статусом 'available'
    - Можно фильтровать по номеру КП
    - Возвращает список словарей с информацией об остатках
    
    Аргументы:
        kp_id: номер КП для фильтрации (None = все КП)
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией об остатках
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        if kp_id:
            cur.execute('''
                SELECT 
                    pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                    pr.length_m, pr.qty, pr.status, pr.created_date,
                    pr.production_day, ko.customer_name
                FROM plate_rests pr
                LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
                WHERE pr.status = 'available' AND pr.kp_id = ?
                ORDER BY pr.created_date DESC, pr.rest_width_mm DESC
            ''', (kp_id,))
        else:
            cur.execute('''
                SELECT 
                    pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                    pr.length_m, pr.qty, pr.status, pr.created_date,
                    pr.production_day, ko.customer_name
                FROM plate_rests pr
                LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
                WHERE pr.status = 'available'
                ORDER BY pr.created_date DESC, pr.rest_width_mm DESC
            ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def mark_rest_as_used(
    rest_id: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Помечает остаток как использованный.
    
    Простыми словами:
    - Когда остаток используется во вторичном резе
    - Его статус меняется на 'used'
    - Он больше не будет показываться как доступный
    
    Аргументы:
        rest_id: ID остатка в БД
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'used', used_date = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} помечен как использованный")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже использован")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при обновлении остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def complete_plate_rest(
    rest_id: int,
    production_day: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Помечает остаток как выполненный (произведённый).
    
    Простыми словами:
    - Когда остаток изготавливается как отдельная плита
    - Его статус меняется на 'completed'
    - Записывается день производства
    
    Аргументы:
        rest_id: ID остатка в БД
        production_day: номер дня производства
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'completed', used_date = ?, production_day = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, production_day, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} выполнен (день {production_day})")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже обработан")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при завершении остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def discard_plate_rest(
    rest_id: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Списывает остаток (брак или утилизация).
    
    Простыми словами:
    - Когда остаток больше не нужен (брак, повреждение)
    - Его статус меняется на 'discarded'
    - Остаток удаляется из доступных
    
    Аргументы:
        rest_id: ID остатка в БД
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'discarded', used_date = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} списан")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже обработан")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при списании остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def get_all_plate_rests(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все остатки (для статистики и отчётов).
    
    Возвращает:
        Список словарей с информацией обо всех остатках
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                pr.length_m, pr.qty, pr.status, pr.created_date,
                pr.used_date, pr.production_day, ko.customer_name
            FROM plate_rests pr
            LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
            ORDER BY pr.created_date DESC, pr.status, pr.rest_width_mm DESC
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def check_and_update_kp_completion(kp_id: int, db_path: str = DEFAULT_DB) -> bool:
    """
    Проверяет, все ли плиты КП выполнены.
    Если да — меняет статус КП на "выполнено".
    
    Простыми словами:
    - Смотрит, остались ли ещё плиты в kp_plates для данного КП
    - Если плит не осталось (все выполнены) — ставит статус "выполнено"
    
    Аргументы:
        kp_id: номер КП для проверки
        db_path: путь к базе данных
    
    Возвращает:
        True если КП полностью выполнен, False если ещё есть плиты
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        # Считаем оставшиеся плиты в КП
        cur.execute('SELECT SUM(qty) FROM kp_plates WHERE kp_id = ?', (kp_id,))
        result = cur.fetchone()
        remaining = result[0] if result[0] else 0
        
        if remaining == 0:
            # Все плиты выполнены — обновляем статус
            cur.execute('''
                UPDATE kp_meta SET status = 'выполнено' WHERE kp_id = ?
            ''', (kp_id,))
            conn.commit()
            print(f"[DB] 🎉 КП #{kp_id} полностью выполнен! Статус обновлён.")
            return True
        
        return False
        
    finally:
        conn.close()


def get_remaining_plates_for_kp(kp_id: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список оставшихся (невыполненных) плит для КП.
    
    Простыми словами:
    - Возвращает все плиты, которые ещё не выполнены для данного КП
    
    Аргументы:
        kp_id: номер КП
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT * FROM kp_plates 
            WHERE kp_id = ? AND qty > 0
            ORDER BY position_number
        ''', (kp_id,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_completed_plates_for_kp(kp_id: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список выполненных плит для КП.
    
    Простыми словами:
    - Возвращает все плиты, которые уже выполнены для данного КП
    
    Аргументы:
        kp_id: номер КП
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о выполненных плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT * FROM completed_plates 
            WHERE kp_id = ?
            ORDER BY completed_date, production_day
        ''', (kp_id,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_completed_plates_stats(db_path: str = DEFAULT_DB) -> Dict:
    """
    Получает статистику по выполненным плитам.
    
    Простыми словами:
    - Считает общее количество выполненных плит
    - Считает сколько КП затронуто
    - Возвращает сводку
    
    Возвращает:
        Словарь со статистикой
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        
        # Общее количество записей
        cur.execute('SELECT COUNT(*) FROM completed_plates')
        total_records = cur.fetchone()[0]
        
        # Общее количество плит (сумма qty)
        cur.execute('SELECT SUM(qty) FROM completed_plates')
        result = cur.fetchone()
        total_qty = result[0] if result[0] else 0
        
        # Количество уникальных КП
        cur.execute('SELECT COUNT(DISTINCT kp_id) FROM completed_plates')
        kp_count = cur.fetchone()[0]
        
        # Количество дней производства
        cur.execute('SELECT COUNT(DISTINCT production_day) FROM completed_plates')
        days_count = cur.fetchone()[0]
        
        return {
            'total_records': total_records,
            'total_plates': total_qty,
            'kp_count': kp_count,
            'days_count': days_count
        }
        
    finally:
        conn.close()


def get_completed_plates_by_day(production_day: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все выполненные плиты за конкретный день производства.
    
    Аргументы:
        production_day: номер дня производства
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT cp.*, ko.customer_name
            FROM completed_plates cp
            LEFT JOIN KP_offers ko ON cp.kp_id = ko.kp_id
            WHERE cp.production_day = ?
            ORDER BY cp.kp_id, cp.plate_name
        ''', (production_day,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_all_plates_in_production(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все плиты в производстве (из таблицы kp_plates).
    
    Простыми словами:
    - Возвращает все плиты, которые ещё не выполнены
    - Объединяет с информацией о КП (клиент, дата)
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                p.id,
                p.kp_id,
                p.position_number,
                p.plate_name,
                p.length_m,
                p.width_m,
                p.load_class,
                p.qty,
                p.unit_weight,
                p.total_weight,
                p.discounted_price,
                ko.customer_name,
                ko.execution_terms
            FROM kp_plates p
            LEFT JOIN KP_offers ko ON p.kp_id = ko.kp_id
            LEFT JOIN kp_meta m ON p.kp_id = m.kp_id
            WHERE p.qty > 0 AND (m.status IS NULL OR m.status = 'в работе')
            ORDER BY p.kp_id, p.position_number
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_all_completed_plates(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все выполненные плиты (из таблицы completed_plates).
    
    Простыми словами:
    - Возвращает все плиты, которые уже выполнены
    - Объединяет с информацией о КП (клиент)
    
    Возвращает:
        Список словарей с информацией о выполненных плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                cp.id,
                cp.kp_id,
                cp.plate_name,
                cp.length_m,
                cp.width_m,
                cp.load_class,
                cp.qty,
                cp.completed_date,
                cp.production_day,
                ko.customer_name,
                ko.execution_terms
            FROM completed_plates cp
            LEFT JOIN KP_offers ko ON cp.kp_id = ko.kp_id
            ORDER BY cp.completed_date DESC, cp.kp_id, cp.plate_name
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_next_kp_number(db_path: str = DEFAULT_DB) -> int:
    """
    Получает следующий свободный номер КП из базы данных.
    
    Простыми словами:
    - Смотрит максимальный номер КП в таблице KP_offers
    - Возвращает следующий номер (max + 1)
    - Если таблица пуста, возвращает 1
    
    Возвращает:
        Следующий номер КП для нового коммерческого предложения
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        cur.execute('SELECT MAX(kp_id) FROM KP_offers')
        result = cur.fetchone()
        
        max_id = result[0] if result[0] is not None else 0
        next_id = max_id + 1
        
        print(f"[DB] Следующий номер КП: {next_id}")
        return next_id
    
    finally:
        conn.close()


# ==================== ФУНКЦИИ ДЛЯ РАБОТЫ С МЕНЕДЖЕРАМИ ====================

def add_manager(
    fio: str,
    contact_number: str,
    email: str,
    db_path: str = DEFAULT_DB
) -> int:
    """
    Добавляет нового менеджера в базу данных.
    
    Простыми словами:
    - Сохраняет информацию о менеджере (ФИО, телефон, email)
    - Email должен быть уникальным (нельзя добавить двух менеджеров с одинаковым email)
    - Возвращает ID созданного менеджера
    
    Аргументы:
        fio: полное имя менеджера (например: "Иванов Иван Иванович")
        contact_number: контактный номер телефона (например: "79621860029")
        email: email адрес (например: "ivanov@example.ru")
        db_path: путь к базе данных
    
    Возвращает:
        ID созданного менеджера или 0 при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        cur.execute('''
            INSERT INTO managers (fio, contact_number, email)
            VALUES (?, ?, ?)
        ''', (fio, contact_number, email))
        
        manager_id = cur.lastrowid
        conn.commit()
        print(f"[DB] ✅ Менеджер добавлен: {fio} (ID: {manager_id})")
        return manager_id
        
    except sqlite3.IntegrityError:
        print(f"[DB] ⚠️ Менеджер с email {email} уже существует")
        conn.rollback()
        return 0
    except Exception as e:
        print(f"[DB] ❌ Ошибка при добавлении менеджера: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def get_all_managers(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список всех менеджеров.
    
    Простыми словами:
    - Возвращает всех менеджеров из базы данных
    - Каждый менеджер представлен словарём с полями: id, fio, contact_number, email
    
    Возвращает:
        Список словарей с информацией о менеджерах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute('SELECT * FROM managers ORDER BY fio')
        return [dict(row) for row in cur.fetchall()]
    
    finally:
        conn.close()


def get_manager_by_id(manager_id: int, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """
    Получает информацию о менеджере по ID.
    
    Простыми словами:
    - Ищет менеджера по его порядковому номеру
    - Возвращает информацию о нём или None, если не найден
    
    Аргументы:
        manager_id: порядковый номер менеджера
        db_path: путь к базе данных
    
    Возвращает:
        Словарь с информацией о менеджере или None
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute('SELECT * FROM managers WHERE id = ?', (manager_id,))
        row = cur.fetchone()
        return dict(row) if row else None
    
    finally:
        conn.close()


def get_manager_by_email(email: str, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """
    Получает информацию о менеджере по email.
    
    Простыми словами:
    - Ищет менеджера по его email адресу
    - Возвращает информацию о нём или None, если не найден
    
    Аргументы:
        email: email адрес менеджера
        db_path: путь к базе данных
    
    Возвращает:
        Словарь с информацией о менеджере или None
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute('SELECT * FROM managers WHERE email = ?', (email,))
        row = cur.fetchone()
        return dict(row) if row else None
    
    finally:
        conn.close()


def update_manager(
    manager_id: int,
    fio: str = None,
    contact_number: str = None,
    email: str = None,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Обновляет информацию о менеджере.
    
    Простыми словами:
    - Меняет данные менеджера (можно обновить только нужные поля)
    - Если передать None для какого-то поля, оно не изменится
    
    Аргументы:
        manager_id: порядковый номер менеджера
        fio: новое ФИО (опционально)
        contact_number: новый контактный номер (опционально)
        email: новый email (опционально)
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False если менеджер не найден
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        
        # Формируем список обновлений
        updates = []
        values = []
        
        if fio is not None:
            updates.append('fio = ?')
            values.append(fio)
        if contact_number is not None:
            updates.append('contact_number = ?')
            values.append(contact_number)
        if email is not None:
            updates.append('email = ?')
            values.append(email)
        
        if not updates:
            return False
        
        values.append(manager_id)
        query = f'UPDATE managers SET {", ".join(updates)} WHERE id = ?'
        
        cur.execute(query, values)
        conn.commit()
        
        if cur.rowcount > 0:
            print(f"[DB] ✅ Менеджер #{manager_id} обновлён")
            return True
        else:
            print(f"[DB] ⚠️ Менеджер #{manager_id} не найден")
            return False
    
    except sqlite3.IntegrityError:
        print(f"[DB] ⚠️ Менеджер с таким email уже существует")
        conn.rollback()
        return False
    except Exception as e:
        print(f"[DB] ❌ Ошибка при обновлении менеджера: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def delete_manager(manager_id: int, db_path: str = DEFAULT_DB) -> bool:
    """
    Удаляет менеджера из базы данных.
    
    Простыми словами:
    - Удаляет менеджера по его порядковому номеру
    - Это необратимая операция
    
    Аргументы:
        manager_id: порядковый номер менеджера для удаления
        db_path: путь к базе данных
    
    Возвращает:
        True если менеджер был найден и удалён, False если не найден
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        cur.execute('DELETE FROM managers WHERE id = ?', (manager_id,))
        conn.commit()
        
        if cur.rowcount > 0:
            print(f"[DB] ✅ Менеджер #{manager_id} удалён")
            return True
        else:
            print(f"[DB] ⚠️ Менеджер #{manager_id} не найден")
            return False
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при удалении менеджера: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def init_default_managers(db_path: str = DEFAULT_DB) -> int:
    """
    Добавляет менеджеров по умолчанию в базу данных.
    
    Простыми словами:
    - Добавляет список менеджеров из вашей таблицы
    - Если менеджер с таким email уже есть, пропускает его
    - Это удобно для первоначальной загрузки данных
    
    Возвращает:
        Количество успешно добавленных менеджеров
    """
    default_managers = [
        ('Булдаков Александр Алексеевич', '79621860029', 'buldakov@gbkstart.ru'),
        ('Зубов Алексей Юрьевич', '79621872265', 'zubov@gbkstart.ru'),
        ('Дургина Ольга Владимировна', '79066395000', 'zakaz@gbkstart.ru'),
        ('Шишов Александр Васильевич', '79206405585', 'shishov@gbkstart.ru'),
        ('Кудигин Никита Валерьевич', '79607428972', 'kuligin@gbkstart.ru'),
    ]
    
    init_schema(db_path)
    added_count = 0
    
    for fio, contact_number, email in default_managers:
        manager_id = add_manager(fio, contact_number, email, db_path)
        if manager_id > 0:
            added_count += 1
    
    print(f"[DB] ✅ Добавлено менеджеров: {added_count} из {len(default_managers)}")
    return added_count


def get_all_kp_list(db_path: str = DEFAULT_DB) -> Dict[str, List[Dict]]:
    """
    Получает все КП, разделенные по статусам.
    
    Простыми словами:
    - Возвращает все КП из базы данных
    - Группирует их по статусам: "в архиве", "в работе", "выполнено"
    - Сортирует по номеру КП (от меньшего к большему)
    
    Возвращает:
        Словарь со списками КП по статусам:
        {
            'archived': [...],      # КП со статусом "в архиве"
            'in_production': [...], # КП со статусом "в работе"
            'completed': [...]      # КП со статусом "выполнено"
        }
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Получаем все КП с их статусами
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
                ko.delivery_conditions,
                ko.payment_conditions,
                ko.execution_terms,
                m.status
            FROM KP_offers ko
            LEFT JOIN kp_meta m ON ko.kp_id = m.kp_id
            ORDER BY ko.kp_id ASC
        ''')
        
        all_kp = [dict(row) for row in cur.fetchall()]
        
        # Группируем по статусам
        result = {
            'archived': [],
            'in_production': [],
            'completed': []
        }
        
        for kp in all_kp:
            status = kp.get('status', 'в работе')  # По умолчанию "в работе"
            
            if status == 'в архиве':
                result['archived'].append(kp)
            elif status == 'в работе':
                result['in_production'].append(kp)
            elif status == 'выполнено':
                result['completed'].append(kp)
        
        return result
        
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
