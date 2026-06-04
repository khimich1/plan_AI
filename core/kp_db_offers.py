#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Commercial offers (KP) persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import sqlite3
import traceback
from typing import Dict, List, Optional

from core.destructive_db_guard import require_destructive_db_reset
from core.kp_db_common import DEFAULT_DB, _connect


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
    status: str = "в работе",
    logistics_cost: float = 0.0,
    db_path: str = DEFAULT_DB,
) -> int:
    """Сохраняет КП в базу (backward-compatible facade).

    Оркестрация — :class:`core.kp_persistence_service.KpPersistenceService`.
    """
    from core.kp_persistence_service import KpPersistenceService

    return KpPersistenceService.save_kp_to_db(
        creation_date,
        order_data,
        xlsx_file_path,
        customer_name,
        manager_name,
        discount_percent,
        delivery_conditions,
        payment_conditions,
        execution_terms,
        status,
        logistics_cost,
        db_path,
    )


def update_kp_discount(kp_id: int, new_discount: float, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет процент скидки для КП, пересчитывает итоги и discounted_price по плитам.
    
    Простыми словами:
    - Берёт unit_price из kp_plates (или восстанавливает из discounted_price и текущей скидки)
    - Считает новые subtotal/vat/total через calculate_total_cost
    - Обновляет KP_offers и discounted_price (и при необходимости unit_price) в kp_plates
    """
    if not (0 <= new_discount <= 100):
        return False
    try:
        from core.commercial_offer_xlsx import calculate_total_cost
    except ImportError:
        return False
    
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        cur.execute(
            'SELECT discount_percent, COALESCE(logistics_cost, 0) FROM KP_offers WHERE kp_id = ?',
            (kp_id,),
        )
        row = cur.fetchone()
        if not row:
            return False
        current_discount = row[0] or 0.0
        logistics_saved = max(0.0, float(row[1] or 0.0))
        
        cur.execute(
            'SELECT id, plate_name, length_m, width_m, load_class, qty, unit_weight, total_weight, discounted_price, unit_price, COALESCE(concrete_grade, "") AS concrete_grade FROM kp_plates WHERE kp_id = ? ORDER BY position_number',
            (kp_id,)
        )
        plates = cur.fetchall()
        if not plates:
            return False
        
        order_data = []
        for p in plates:
            (
                pid,
                plate_name,
                length_m,
                width_m,
                load_class,
                qty,
                unit_weight,
                total_weight,
                discounted_price,
                unit_price_col,
                concrete_grade_col,
            ) = p
            if unit_price_col is not None and unit_price_col > 0:
                unit_price = float(unit_price_col)
            else:
                # Восстанавливаем из discounted_price и текущей скидки
                factor = 1.0 - (current_discount / 100.0)
                if factor <= 0:
                    factor = 1.0
                unit_price = (discounted_price or 0) / factor
            weight = total_weight if total_weight is not None and total_weight > 0 else (unit_weight or 0) * (qty or 0)
            order_data.append({
                'name': plate_name or '',
                'length_m': length_m or 0,
                'width_m': width_m or 0,
                'qty': qty or 0,
                'load_class': load_class or 800,
                'unit_price': unit_price,
                'weight': weight,
                'concrete_grade': (concrete_grade_col or '').strip() or None,
            })
        
        totals = calculate_total_cost(order_data, new_discount, logistics_cost=logistics_saved)
        subtotal = totals['subtotal']
        vat_amount = totals['vat_amount']
        total_amount = totals['total_with_vat']
        
        cur.execute('''
            UPDATE KP_offers SET discount_percent = ?, subtotal = ?, vat_amount = ?, total_amount = ?
            WHERE kp_id = ?
        ''', (new_discount, subtotal, vat_amount, total_amount, kp_id))
        
        for p in plates:
            (
                pid,
                plate_name,
                length_m,
                width_m,
                load_class,
                qty,
                unit_weight,
                total_weight,
                discounted_price,
                unit_price_col,
                _conc2,
            ) = p
            if unit_price_col is not None and unit_price_col > 0:
                up = float(unit_price_col)
            else:
                factor = 1.0 - (current_discount / 100.0)
                if factor <= 0:
                    factor = 1.0
                up = (discounted_price or 0) / factor
            new_disc_price = up * (1.0 - new_discount / 100.0)
            cur.execute(
                'UPDATE kp_plates SET discounted_price = ?, unit_price = ? WHERE id = ?',
                (new_disc_price, up, pid)
            )
        
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        conn.close()


def update_kp_logistics_cost(kp_id: int, logistics_cost: float, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет стоимость одного рейса (поле KP_offers.logistics_cost) и пересчитывает
    суммы KP_offers согласно calculate_total_cost (как при генерации PDF/XLSX КП).
    Цены по плитам не меняются.
    """
    trip = max(0.0, float(logistics_cost or 0.0))
    try:
        from core.commercial_offer_xlsx import calculate_total_cost
    except ImportError:
        return False

    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()

        cur.execute(
            "SELECT discount_percent FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        )
        row = cur.fetchone()
        if not row:
            return False
        current_discount = float(row[0] or 0.0)

        cur.execute(
            "SELECT id, plate_name, length_m, width_m, load_class, qty, unit_weight, total_weight, "
            "discounted_price, unit_price FROM kp_plates WHERE kp_id = ? ORDER BY position_number",
            (kp_id,),
        )
        plates = cur.fetchall()
        if not plates:
            return False

        order_data: list[dict] = []
        for p in plates:
            pid, plate_name, length_m, width_m, load_class, qty, unit_weight, total_weight, discounted_price, unit_price_col = p
            if unit_price_col is not None and unit_price_col > 0:
                unit_price = float(unit_price_col)
            else:
                factor = 1.0 - (current_discount / 100.0)
                if factor <= 0:
                    factor = 1.0
                unit_price = (discounted_price or 0) / factor
            weight = total_weight if total_weight is not None and total_weight > 0 else (unit_weight or 0) * (qty or 0)
            order_data.append(
                {
                    "name": plate_name or "",
                    "length_m": length_m or 0,
                    "width_m": width_m or 0,
                    "qty": qty or 0,
                    "load_class": load_class or 800,
                    "unit_price": unit_price,
                    "weight": weight,
                }
            )

        totals = calculate_total_cost(order_data, current_discount, logistics_cost=trip)
        subtotal = totals["subtotal"]
        vat_amount = totals["vat_amount"]
        total_amount = totals["total_with_vat"]

        cur.execute(
            """
            UPDATE KP_offers
            SET logistics_cost = ?, subtotal = ?, vat_amount = ?, total_amount = ?
            WHERE kp_id = ?
            """,
            (trip, subtotal, vat_amount, total_amount, kp_id),
        )

        conn.commit()
        return True
    except Exception:
        return False
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
    from core.kp_file_paths import resolve_kp_xlsx_path_for_read

    resolved_xlsx = resolve_kp_xlsx_path_for_read(xlsx_file_path)
    if resolved_xlsx is None:
        return False

    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        with open(resolved_xlsx, 'rb') as f:
            xlsx_blob = f.read()

        safe_path = str(resolved_xlsx)
        cur = conn.cursor()
        # Проверяем, есть ли уже запись
        cur.execute('SELECT id FROM kp_files WHERE kp_id = ?', (kp_id,))
        exists = cur.fetchone()

        if exists:
            # Обновляем существующую запись
            cur.execute(
                '''
                UPDATE kp_files
                SET xlsx_file = ?, file_path = ?
                WHERE kp_id = ?
                ''',
                (xlsx_blob, safe_path, kp_id),
            )
        else:
            # Создаём новую запись
            cur.execute(
                '''
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
                ''',
                (kp_id, xlsx_blob, safe_path),
            )
        
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
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        cur.execute('SELECT xlsx_file FROM kp_files WHERE kp_id = ?', (kp_id,))
        row = cur.fetchone()
        
        if not row or not row[0]:
            return None
        
        xlsx_data = row[0]

        if output_path:
            from core.kp_file_paths import resolve_kp_xlsx_path_for_write

            safe_path = resolve_kp_xlsx_path_for_write(output_path)
            if safe_path is None:
                return None
            safe_path.parent.mkdir(parents=True, exist_ok=True)
            with open(safe_path, 'wb') as f:
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
    - Удаляет журнал переходов статусов ``plate_status_log``
    - Удаляет ВСЕ остатки от резки
    - Сбрасывает счётчики AUTOINCREMENT
    - Это как полностью очистить завод от всех заказов и начать с нуля
    
    ⚠️ ВНИМАНИЕ: Это необратимая операция!
    
    Возвращает:
        Словарь с количеством удалённых записей из каждой таблицы
    """
    require_destructive_db_reset()
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

        cur.execute('SELECT COUNT(*) FROM plate_status_log')
        status_log_count = cur.fetchone()[0]
        
        # Удаляем все записи из всех таблиц
        # Порядок важен: сначала зависимые таблицы, потом основную
        cur.execute('DELETE FROM plate_status_log')
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
                          'plate_rests', 'kp_files', 'kp_meta',
                          'plate_status_log')
        ''')
        
        conn.commit()
        
        result = {
            'kp_offers': kp_count,
            'kp_plates': plates_count,
            'completed_plates': completed_count,
            'plate_rests': rests_count,
            'kp_files': files_count,
            'kp_meta': meta_count,
            'plate_status_log': status_log_count,
            'total': (
                kp_count
                + plates_count
                + completed_count
                + rests_count
                + files_count
                + meta_count
                + status_log_count
            ),
        }
        
        print(f"[DB] ✅ ПОЛНАЯ ОЧИСТКА ВСЕХ ДАННЫХ О ПЛИТАХ:")
        print(f"  - Удалено КП: {kp_count}")
        print(f"  - Удалено плит в работе: {plates_count}")
        print(f"  - Удалено выполненных плит: {completed_count}")
        print(f"  - Удалено остатков: {rests_count}")
        print(f"  - Удалено файлов: {files_count}")
        print(f"  - Удалено метаданных: {meta_count}")
        print(f"  - Удалено записей журнала статусов: {status_log_count}")
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
    require_destructive_db_reset()
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


def _escape_sql_like(value: str) -> str:
    """Экранирует спецсимволы LIKE (% и _) в пользовательском вводе."""
    return value.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def search_kp_by_customer_name(
    name: str,
    limit: int = 50,
    db_path: str = DEFAULT_DB,
) -> tuple[List[Dict], int]:
    """
    Ищет КП по частичному совпадению имени заказчика (глобально, все статусы).

    Returns:
        (список строк для списка архива, общее число совпадений)
    """
    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        escaped = _escape_sql_like(name.strip())
        pattern = f"%{escaped}%"
        fetch_limit = limit + 1

        base_select = """
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
                m.status
            FROM KP_offers ko
            LEFT JOIN kp_meta m ON ko.kp_id = m.kp_id
            WHERE casefold(ko.customer_name) LIKE casefold(?) ESCAPE '\\'
        """

        cur.execute(
            f"{base_select} ORDER BY ko.kp_id DESC LIMIT ?",
            (pattern, fetch_limit),
        )
        rows = [dict(row) for row in cur.fetchall()]

        if len(rows) > limit:
            cur.execute(
                f"SELECT COUNT(*) AS cnt FROM KP_offers ko WHERE casefold(ko.customer_name) LIKE casefold(?) ESCAPE '\\'",
                (pattern,),
            )
            total = int(cur.fetchone()["cnt"])
            return rows[:limit], total

        return rows, len(rows)
    finally:
        conn.close()


def get_kp_completion_percentage(kp_id: int, db_path: str = DEFAULT_DB) -> Dict:
    """
    Подсчитывает процент выполнения КП.
    
    Простыми словами:
    - Считает сколько плит уже выполнено
    - Считает сколько всего плит в КП
    - Возвращает процент выполнения
    
    Args:
        kp_id: номер КП
        db_path: путь к базе данных
        
    Returns:
        Словарь с информацией:
        {
            'total_plates': int,         # Всего плит
            'completed_plates': int,     # Выполненные плиты  
            'in_production': int,        # В производстве
            'percentage': float          # Процент выполнения (0-100)
        }
    """
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Считаем плиты в производстве (из kp_plates)
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as qty
            FROM kp_plates
            WHERE kp_id = ?
        ''', (kp_id,))
        in_production = cur.fetchone()['qty']
        
        # Считаем выполненные плиты (из completed_plates)
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as qty
            FROM completed_plates
            WHERE kp_id = ?
        ''', (kp_id,))
        completed = cur.fetchone()['qty']
        
        # Всего плит
        total = in_production + completed
        
        # Процент выполнения
        if total > 0:
            percentage = (completed / total) * 100
        else:
            percentage = 0.0
        
        return {
            'total_plates': total,
            'completed_plates': completed,
            'in_production': in_production,
            'percentage': round(percentage, 1)
        }
        
    finally:
        conn.close()


def get_kp_plates_in_plan_percentage(kp_id: int, db_path: str = DEFAULT_DB) -> Dict:
    """
    Подсчитывает, какой процент плит КП уже в плане производства.

    Считает сумму qty по kp_plates со статусом 'в производстве' или 'в плане' (база),
    и сумму qty со статусом 'в плане'; возвращает процент (in_plan / total * 100).

    Args:
        kp_id: номер КП
        db_path: путь к базе данных

    Returns:
        Словарь: 'total_plates', 'in_plan', 'percentage' (0-100)
    """
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as total
            FROM kp_plates
            WHERE kp_id = ? AND status IN ('в производстве', 'в плане')
        ''', (kp_id,))
        total = cur.fetchone()[0]
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as in_plan
            FROM kp_plates
            WHERE kp_id = ? AND status = 'в плане'
        ''', (kp_id,))
        in_plan = cur.fetchone()[0]
        percentage = (in_plan / total * 100) if total > 0 else 0.0
        return {
            'total_plates': total,
            'in_plan': in_plan,
            'percentage': round(percentage, 1)
        }
    finally:
        conn.close()


def get_kp_total_length(kp_id: int, db_path: str = DEFAULT_DB) -> float:
    """
    Суммарная длина плит КП в метрах: SUM(COALESCE(length_m, 0) * qty)
    по строкам kp_plates со статусом 'в производстве' или 'в плане'.
    """
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('''
            SELECT COALESCE(SUM(COALESCE(length_m, 0) * qty), 0.0)
            FROM kp_plates
            WHERE kp_id = ? AND status IN ('в производстве', 'в плане')
        ''', (kp_id,))
        return float(cur.fetchone()[0])
    finally:
        conn.close()


def update_kp_execution_date(kp_id: int, new_date: str, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет дату выполнения для КП.
    
    Простыми словами:
    - Меняет срок выполнения для КП
    - Это влияет на все плиты в этом КП
    
    Args:
        kp_id: номер КП
        new_date: новая дата в формате "DD.MM.YYYY"
        db_path: путь к базе данных
        
    Returns:
        True при успехе, False при ошибке
    """
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        
        # Обновляем дату выполнения
        cur.execute('''
            UPDATE KP_offers 
            SET execution_terms = ? 
            WHERE kp_id = ?
        ''', (new_date, kp_id))
        
        conn.commit()
        
        # Проверяем, что обновление произошло
        return cur.rowcount > 0
        
    except Exception as e:
        conn.rollback()
        return False
        
    finally:
        conn.close()
