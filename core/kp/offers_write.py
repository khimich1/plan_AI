"""Write/admin SQL for commercial offers (KP) — extracted from kp_db_offers (A5 slice 5)."""

from __future__ import annotations

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
    owner_user_id: int | None = None,
    product_type: str = "plates",
    db_path: str = DEFAULT_DB,
) -> int:
    """Сохраняет КП в базу.

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
        owner_user_id,
        product_type,
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
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()

        cur.execute(
            "SELECT discount_percent, COALESCE(logistics_cost, 0) FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        )
        row = cur.fetchone()
        if not row:
            return False
        current_discount = row[0] or 0.0
        logistics_saved = max(0.0, float(row[1] or 0.0))

        cur.execute(
            "SELECT id, plate_name, length_m, width_m, load_class, qty, unit_weight, total_weight, discounted_price, unit_price, COALESCE(concrete_grade, '') AS concrete_grade FROM kp_plates WHERE kp_id = ? ORDER BY position_number",
            (kp_id,),
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
                    "concrete_grade": (concrete_grade_col or "").strip() or None,
                }
            )

        totals = calculate_total_cost(order_data, new_discount, logistics_cost=logistics_saved)
        subtotal = totals["subtotal"]
        vat_amount = totals["vat_amount"]
        total_amount = totals["total_with_vat"]

        cur.execute(
            """
            UPDATE KP_offers SET discount_percent = ?, subtotal = ?, vat_amount = ?, total_amount = ?
            WHERE kp_id = ?
        """,
            (new_discount, subtotal, vat_amount, total_amount, kp_id),
        )

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
                "UPDATE kp_plates SET discounted_price = ?, unit_price = ? WHERE id = ?",
                (new_disc_price, up, pid),
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
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE kp_meta
            SET status = ?
            WHERE kp_id = ?
        """,
            (new_status, kp_id),
        )

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
        conn.execute("PRAGMA foreign_keys = ON")
        with open(resolved_xlsx, "rb") as f:
            xlsx_blob = f.read()

        safe_path = str(resolved_xlsx)
        cur = conn.cursor()
        cur.execute("SELECT id FROM kp_files WHERE kp_id = ?", (kp_id,))
        exists = cur.fetchone()

        if exists:
            cur.execute(
                """
                UPDATE kp_files
                SET xlsx_file = ?, file_path = ?
                WHERE kp_id = ?
                """,
                (xlsx_blob, safe_path, kp_id),
            )
        else:
            cur.execute(
                """
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
                """,
                (kp_id, xlsx_blob, safe_path),
            )

        conn.commit()
        return True

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
        conn.execute("PRAGMA foreign_keys = ON")

        cur = conn.cursor()

        cur.execute("SELECT kp_id FROM KP_offers WHERE kp_id = ?", (kp_id,))
        if not cur.fetchone():
            return False

        cur.execute("DELETE FROM KP_offers WHERE kp_id = ?", (kp_id,))

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

    ⚠️ ВНИМАНИЕ: Это необратимая операция!

    Возвращает:
        Словарь с количеством удалённых записей из каждой таблицы
    """
    require_destructive_db_reset()
    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")

        cur = conn.cursor()

        cur.execute("SELECT COUNT(*) FROM KP_offers")
        kp_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM kp_plates")
        plates_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM completed_plates")
        completed_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM plate_rests")
        rests_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM kp_files")
        files_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM kp_meta")
        meta_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM plate_status_log")
        status_log_count = cur.fetchone()[0]

        cur.execute("DELETE FROM plate_status_log")
        cur.execute("DELETE FROM kp_plates")
        cur.execute("DELETE FROM completed_plates")
        cur.execute("DELETE FROM plate_rests")
        cur.execute("DELETE FROM kp_files")
        cur.execute("DELETE FROM kp_meta")
        cur.execute("DELETE FROM KP_offers")

        cur.execute(
            """
            DELETE FROM sqlite_sequence
            WHERE name IN ('KP_offers', 'kp_plates', 'completed_plates',
                          'plate_rests', 'kp_files', 'kp_meta',
                          'plate_status_log')
        """
        )

        conn.commit()

        result = {
            "kp_offers": kp_count,
            "kp_plates": plates_count,
            "completed_plates": completed_count,
            "plate_rests": rests_count,
            "kp_files": files_count,
            "kp_meta": meta_count,
            "plate_status_log": status_log_count,
            "total": (
                kp_count
                + plates_count
                + completed_count
                + rests_count
                + files_count
                + meta_count
                + status_log_count
            ),
        }

        print("[DB] ✅ ПОЛНАЯ ОЧИСТКА ВСЕХ ДАННЫХ О ПЛИТАХ:")
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
        traceback.print_exc()
        conn.rollback()
        raise

    finally:
        conn.close()


def clear_all_kp(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """
    Полностью очищает все таблицы с КП из базы данных.

    Возвращает:
        Словарь с количеством удалённых записей из каждой таблицы
    """
    require_destructive_db_reset()
    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")

        cur = conn.cursor()

        cur.execute("SELECT COUNT(*) FROM KP_offers")
        kp_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM kp_plates")
        plates_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM kp_files")
        files_count = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM kp_meta")
        meta_count = cur.fetchone()[0]

        cur.execute("DELETE FROM kp_plates")
        cur.execute("DELETE FROM kp_files")
        cur.execute("DELETE FROM kp_meta")
        cur.execute("DELETE FROM KP_offers")

        cur.execute(
            "DELETE FROM sqlite_sequence WHERE name IN (?, ?, ?, ?)",
            ("KP_offers", "kp_plates", "kp_files", "kp_meta"),
        )

        conn.commit()

        result = {
            "kp_offers": kp_count,
            "kp_plates": plates_count,
            "kp_files": files_count,
            "kp_meta": meta_count,
        }

        print("[DB] ✅ Полная очистка БД завершена:")
        print(f"  - Удалено КП: {kp_count}")
        print(f"  - Удалено записей плит: {plates_count}")
        print(f"  - Удалено файлов: {files_count}")
        print(f"  - Удалено метаданных: {meta_count}")

        return result

    except Exception as e:
        print(f"[DB] ❌ Ошибка при полной очистке БД: {e}")
        traceback.print_exc()
        raise

    finally:
        conn.close()


def update_kp_execution_date(kp_id: int, new_date: str, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет дату выполнения для КП.

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

        cur.execute(
            """
            UPDATE KP_offers
            SET execution_terms = ?
            WHERE kp_id = ?
        """,
            (new_date, kp_id),
        )

        conn.commit()

        return cur.rowcount > 0

    except Exception:
        conn.rollback()
        return False

    finally:
        conn.close()
