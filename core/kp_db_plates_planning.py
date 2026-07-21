#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Plate planning and status transitions (A1 slice)."""

from __future__ import annotations

import sqlite3
from collections import Counter
from typing import Dict, List, Optional, Tuple

from core.kp_db_audit import audit_append
from core.kp_db_common import DEFAULT_DB, _connect
from core.kp_db_plates_common import insert_kp_plate_remainder_row

def mark_plates_as_planned(
    kp_id: int,
    plate_name: str,
    qty_to_plan: int,
    plan_id: str,
    db_path: str = DEFAULT_DB,
    *,
    actor: str | None = None,
    day_number: Optional[int] = None,
) -> Dict[str, object]:
    """
    Помечает плиты как 'в плане' при сохранении плана.

    ИСПРАВЛЕНО: Теперь обрабатывает ВСЕ записи с одинаковым plate_name,
    а не только первую (убран LIMIT 1).

    Простыми словами:
    - Находит ВСЕ плиты по kp_id и plate_name со статусом 'в производстве'
    - Обрабатывает их по очереди, пока не наберется нужное qty_to_plan
    - Если qty_to_plan = всему qty — просто меняет статус на 'в плане'
    - Если qty_to_plan < qty — разбивает запись на две:
      * Одна с qty_to_plan и статусом 'в плане'
      * Вторая с остатком и статусом 'в производстве'

    Аргументы:
        kp_id: номер КП
        plate_name: название плиты (например, "ПБ 60-12-8п")
        qty_to_plan: сколько плит добавить в план
        plan_id: ID плана (для связи и отката)
        db_path: путь к базе данных
        day_number: номер дня производства (P5). Если задан — пишется в
            ``kp_plates.day_number`` и в audit-лог. ``None`` оставляет
            старое поведение (день не зафиксирован — legacy).

    Возвращает:
        Dict с подробным результатом пометки. ``updated_ids`` — id строк
        kp_plates, которые получили статус «в плане» (нужно вызывающему
        слою, чтобы записать ``kp_plate_id`` в plan.json/items).
    """
    conn = _connect(db_path)
    requested_qty = max(int(qty_to_plan or 0), 0)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        # ИСПРАВЛЕНО: Находим ВСЕ плиты со статусом 'в производстве' (без LIMIT 1)
        cur.execute('''
            SELECT id, qty, position_number, length_m, width_m, load_class,
                   unit_weight, total_weight, discounted_price,
                   COALESCE(concrete_grade, '') AS concrete_grade
            FROM kp_plates
            WHERE kp_id = ? AND plate_name = ? AND status = 'в производстве' AND qty > 0
            ORDER BY id
        ''', (kp_id, plate_name))
        
        rows = cur.fetchall()
        if not rows:
            print(f"[DB] ⚠️ Плита не найдена: КП #{kp_id}, {plate_name} (статус 'в производстве')")
            return {
                'success': False,
                'requested_qty': requested_qty,
                'available_qty': 0,
                'processed_count': 0,
                'remaining_unplanned': requested_qty,
                'split_count': 0,
                'updated_ids': [],
                'id_qty_pairs': [],
                'error': "plate_not_found",
            }
        
        # Подсчитываем общее доступное количество
        total_available = sum(row[1] for row in rows)
        
        if qty_to_plan > total_available:
            print(f"[DB] ⚠️ Запрошено {qty_to_plan}, но доступно только {total_available}")
            qty_to_plan = total_available
        
        # Обрабатываем записи по очереди
        remaining_to_plan = qty_to_plan
        processed_count = 0
        split_count = 0
        updated_ids: List[int] = []
        # P5: id_qty_pairs нужен caller'у, чтобы при записи kp_plate_id в plan.json
        # знать, сколько items приходится на каждый plate_id.
        id_qty_pairs: List[Tuple[int, int]] = []

        for row in rows:
            if remaining_to_plan <= 0:
                break
            
            plate_id, current_qty, pos_num, length_m, width_m, load_class, unit_w, total_w, price, conc_grade = row
            
            if current_qty <= remaining_to_plan:
                # Вся запись идет в план
                cur.execute('''
                    UPDATE kp_plates
                    SET status = 'в плане', plan_id = ?, day_number = ?
                    WHERE id = ?
                ''', (plan_id, day_number, plate_id))
                audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=plan_id,
                    day_number=day_number,
                    from_status='в производстве',
                    to_status='в плане',
                    qty=current_qty,
                    reason='planned',
                    actor=actor,
                )
                print(f"[DB] ✅ Плита {plate_name} x{current_qty} помечена как 'в плане' (запись #{plate_id})")
                remaining_to_plan -= current_qty
                processed_count += current_qty
                updated_ids.append(plate_id)
                id_qty_pairs.append((plate_id, current_qty))
            else:
                # Частичная обработка: разбиваем запись
                qty_for_plan = remaining_to_plan
                remaining_in_production = current_qty - qty_for_plan

                # 1. Обновляем существующую запись: уменьшаем qty, помечаем 'в плане'
                cur.execute('''
                    UPDATE kp_plates
                    SET qty = ?, status = 'в плане', plan_id = ?, day_number = ?
                    WHERE id = ?
                ''', (qty_for_plan, plan_id, day_number, plate_id))

                insert_kp_plate_remainder_row(
                    cur,
                    source_plate_id=plate_id,
                    remainder_qty=remaining_in_production,
                    status="в производстве",
                    plan_id=None,
                    day_number=None,
                )
                audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=plan_id,
                    day_number=day_number,
                    from_status='в производстве',
                    to_status='в плане',
                    qty=qty_for_plan,
                    reason='planned',
                    actor=actor,
                )
                
                print(f"[DB] ✅ Плита {plate_name} разбита: {qty_for_plan} в план, {remaining_in_production} осталось (запись #{plate_id})")
                remaining_to_plan = 0
                processed_count += qty_for_plan
                split_count += 1
                updated_ids.append(plate_id)
                id_qty_pairs.append((plate_id, qty_for_plan))
        
        print(f"[DB] ✅ Итого помечено {processed_count} плит '{plate_name}' как 'в плане' (план {plan_id})")
        conn.commit()
        return {
            'success': True,
            'requested_qty': requested_qty,
            'available_qty': total_available,
            'processed_count': processed_count,
            'remaining_unplanned': max(requested_qty - processed_count, 0),
            'split_count': split_count,
            'updated_ids': updated_ids,
            'id_qty_pairs': id_qty_pairs,
            'error': None,
        }
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при пометке плиты как 'в плане': {e}")
        conn.rollback()
        return {
            'success': False,
            'requested_qty': requested_qty,
            'available_qty': 0,
            'processed_count': 0,
            'remaining_unplanned': requested_qty,
            'split_count': 0,
            'updated_ids': [],
            'id_qty_pairs': [],
            'error': str(e),
        }
    
    finally:
        conn.close()


def return_plates_to_production(
    kp_id: int,
    plate_name: str,
    qty: int,
    db_path: str = DEFAULT_DB,
    *,
    actor: str | None = None,
    reason: str = "rejected",
    _external_conn: Optional[sqlite3.Connection] = None,
) -> bool:
    """
    Возвращает плиты обратно в статус 'в производстве'.

    ИСПРАВЛЕНО: Теперь обрабатывает ВСЕ записи с одинаковым plate_name,
    а не только первую (убран LIMIT 1).

    Простыми словами:
    - Используется при браке: бракованные плиты возвращаются в производство
    - Находит ВСЕ плиты со статусом 'в плане' и меняет статус на 'в производстве'
    - Очищает plan_id
    - Обрабатывает нужное количество (qty) по записям

    Аргументы:
        kp_id: номер КП
        plate_name: название плиты
        qty: количество плит для возврата
        db_path: путь к базе данных
        _external_conn: если задано — функция работает в существующей транзакции
            переданного соединения (P0). Не делает commit/rollback и не закрывает conn.
            Все исключения пробрасываются вызывающему слою.

    Возвращает:
        True если успешно, False при ошибке
    """
    own_conn = _external_conn is None
    if own_conn:
        conn = _connect(db_path)
    else:
        conn = _external_conn

    try:
        if own_conn:
            conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()

        # ИСПРАВЛЕНО: Находим ВСЕ плиты со статусом 'в плане' (без LIMIT 1)
        cur.execute('''
            SELECT id, qty
            FROM kp_plates
            WHERE kp_id = ? AND plate_name = ? AND status = 'в плане' AND qty > 0
            ORDER BY id
        ''', (kp_id, plate_name))

        rows = cur.fetchall()
        if not rows:
            # Плита могла уже быть возвращена или не была в плане
            print(f"[DB] ⚠️ Плита для возврата не найдена: КП #{kp_id}, {plate_name}")
            return False
        
        # Обрабатываем записи по очереди
        remaining_to_return = qty
        processed_count = 0
        
        for row in rows:
            if remaining_to_return <= 0:
                break
            
            plate_id, current_qty = row
            
            if current_qty <= remaining_to_return:
                # Вся запись возвращается в производство
                cur.execute('''
                    UPDATE kp_plates
                    SET status = 'в производстве', plan_id = NULL
                    WHERE id = ?
                ''', (plate_id,))
                audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=None,
                    day_number=None,
                    from_status='в плане',
                    to_status='в производстве',
                    qty=current_qty,
                    reason=reason,
                    actor=actor,
                )
                print(f"[DB] ✅ Плита {plate_name} x{current_qty} возвращена в производство (запись #{plate_id})")
                remaining_to_return -= current_qty
                processed_count += current_qty
            else:
                # Частичный возврат: уменьшаем qty в плане, создаём новую запись «в производстве»
                new_qty_in_plan = current_qty - remaining_to_return
                cur.execute('''
                    UPDATE kp_plates
                    SET qty = ?
                    WHERE id = ?
                ''', (new_qty_in_plan, plate_id))
                insert_kp_plate_remainder_row(
                    cur,
                    source_plate_id=plate_id,
                    remainder_qty=remaining_to_return,
                    status="в производстве",
                    plan_id=None,
                    day_number=None,
                )
                audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=None,
                    day_number=None,
                    from_status='в плане',
                    to_status='в производстве',
                    qty=remaining_to_return,
                    reason=reason,
                    actor=actor,
                )
                print(f"[DB] ✅ Частичный возврат: {plate_name} x{remaining_to_return} в производство (осталось в плане: {new_qty_in_plan}, запись #{plate_id})")
                processed_count += remaining_to_return
                remaining_to_return = 0
                break
        
        if own_conn:
            conn.commit()
        print(f"[DB] ✅ Итого возвращено {processed_count} плит '{plate_name}' в производство (КП #{kp_id})")
        return True

    except Exception as e:
        if own_conn:
            print(f"[DB] ❌ Ошибка при возврате плиты в производство: {e}")
            conn.rollback()
            return False
        # Внешняя транзакция: пробрасываем исключение для отката caller'ом
        raise
    
    finally:
        if own_conn:
            conn.close()


def return_plate_rows_for_plan(
    plan_id: str,
    id_qty: Counter[int, int],
    db_path: str = DEFAULT_DB,
    *,
    actor: str | None = None,
    reason: str = "track_removed",
    legacy_identity_qty: Counter[tuple[int, str], int] | None = None,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> dict:
    """
    Возвращает указанные плиты из плана обратно в «в производстве» по kp_plates.id.

    Основной путь — ``id_qty`` (plate_id → qty). Для legacy-планов без kp_plate_id
    можно передать ``legacy_identity_qty`` ((kp_id, plate_name) → qty) с фильтром
    ``plan_id``.

    Возвращает статистику: ``plates_returned``, ``warnings``.
    """
    own_conn = _external_conn is None
    if own_conn:
        conn = _connect(db_path)
    else:
        conn = _external_conn

    plates_returned = 0
    warnings: List[str] = []

    def _return_qty_from_row(
        cur: sqlite3.Cursor,
        *,
        plate_id: int,
        current_qty: int,
        qty_to_return: int,
        kp_id: int,
        plate_name: str,
    ) -> int:
        nonlocal plates_returned
        if qty_to_return <= 0 or current_qty <= 0:
            return 0

        if current_qty <= qty_to_return:
            cur.execute('''
                UPDATE kp_plates
                SET status = 'в производстве', plan_id = NULL
                WHERE id = ?
            ''', (plate_id,))
            audit_append(
                cur,
                plate_id=plate_id,
                kp_id=kp_id,
                plate_name=plate_name,
                plan_id=None,
                day_number=None,
                from_status='в плане',
                to_status='в производстве',
                qty=current_qty,
                reason=reason,
                actor=actor,
            )
            plates_returned += current_qty
            return current_qty

        new_qty_in_plan = current_qty - qty_to_return
        cur.execute('''
            UPDATE kp_plates
            SET qty = ?
            WHERE id = ?
        ''', (new_qty_in_plan, plate_id))
        insert_kp_plate_remainder_row(
            cur,
            source_plate_id=plate_id,
            remainder_qty=qty_to_return,
            status="в производстве",
            plan_id=None,
            day_number=None,
        )
        audit_append(
            cur,
            plate_id=plate_id,
            kp_id=kp_id,
            plate_name=plate_name,
            plan_id=None,
            day_number=None,
            from_status='в плане',
            to_status='в производстве',
            qty=qty_to_return,
            reason=reason,
            actor=actor,
        )
        plates_returned += qty_to_return
        return qty_to_return

    def _fetch_planned_rows_by_identity(
        cur: sqlite3.Cursor,
        kp_id: int,
        plate_name: str,
    ) -> List[Tuple[int, int, int, str]]:
        cur.execute('''
            SELECT id, qty, kp_id, plate_name
            FROM kp_plates
            WHERE kp_id = ? AND plate_name = ? AND plan_id = ?
              AND status = 'в плане' AND qty > 0
            ORDER BY id
        ''', (kp_id, plate_name, plan_id))
        rows = cur.fetchall()
        if rows:
            return rows

        from core import plate_name as _pn
        canon = _pn.canonical(plate_name)
        if not canon:
            return []

        cur.execute('''
            SELECT id, qty, kp_id, plate_name
            FROM kp_plates
            WHERE kp_id = ? AND plan_id = ? AND status = 'в плане' AND qty > 0
            ORDER BY id
        ''', (kp_id, plan_id))
        return [r for r in cur.fetchall() if _pn.canonical(r[3]) == canon]

    try:
        if own_conn:
            conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()

        for plate_id, qty_requested in id_qty.items():
            qty_requested = int(qty_requested or 0)
            if qty_requested <= 0:
                continue

            cur.execute('''
                SELECT id, qty, kp_id, plate_name
                FROM kp_plates
                WHERE id = ? AND plan_id = ? AND status = 'в плане' AND qty > 0
            ''', (plate_id, plan_id))
            row = cur.fetchone()
            if not row:
                warnings.append(
                    f"plate_id={plate_id}: не найдена в плане {plan_id!r} (статус «в плане»)"
                )
                continue

            _pid, current_qty, kp_id, plate_name = row
            if qty_requested > current_qty:
                warnings.append(
                    f"plate_id={plate_id}: запрошено {qty_requested}, доступно {current_qty}"
                )
            _return_qty_from_row(
                cur,
                plate_id=plate_id,
                current_qty=current_qty,
                qty_to_return=min(qty_requested, current_qty),
                kp_id=kp_id,
                plate_name=plate_name,
            )

        if legacy_identity_qty:
            for (kp_id, plate_name), qty_requested in legacy_identity_qty.items():
                qty_requested = int(qty_requested or 0)
                if qty_requested <= 0:
                    continue

                rows = _fetch_planned_rows_by_identity(cur, int(kp_id), plate_name)
                if not rows:
                    warnings.append(
                        f"КП #{kp_id}, {plate_name!r}: не найдена в плане {plan_id!r}"
                    )
                    continue

                available = sum(r[1] for r in rows)
                if qty_requested > available:
                    warnings.append(
                        f"КП #{kp_id}, {plate_name!r}: запрошено {qty_requested}, "
                        f"доступно {available}"
                    )

                remaining = min(qty_requested, available)
                for row in rows:
                    if remaining <= 0:
                        break
                    row_id, row_qty, row_kp_id, row_plate_name = row
                    take = min(remaining, row_qty)
                    returned = _return_qty_from_row(
                        cur,
                        plate_id=row_id,
                        current_qty=row_qty,
                        qty_to_return=take,
                        kp_id=row_kp_id,
                        plate_name=row_plate_name,
                    )
                    remaining -= returned

        if own_conn:
            conn.commit()
        return {"plates_returned": plates_returned, "warnings": warnings}

    except Exception as e:
        if own_conn:
            conn.rollback()
            return {
                "plates_returned": 0,
                "warnings": [f"ошибка возврата плит из плана: {e}"],
            }
        raise

    finally:
        if own_conn:
            conn.close()


def return_plan_plates_to_production(plan_id: str, db_path: str = DEFAULT_DB) -> int:
    """
    Возвращает ВСЕ плиты плана обратно в производство.
    
    Простыми словами:
    - Используется при удалении плана
    - Находит все плиты с данным plan_id
    - Меняет их статус на 'в производстве'
    - Очищает plan_id
    
    Аргументы:
        plan_id: ID плана
        db_path: путь к базе данных
    
    Возвращает:
        Количество возвращённых плит (записей)
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        # Считаем сколько плит вернём
        cur.execute('''
            SELECT COUNT(*), COALESCE(SUM(qty), 0)
            FROM kp_plates
            WHERE plan_id = ?
        ''', (plan_id,))
        
        count_result = cur.fetchone()
        records_count = count_result[0] if count_result else 0
        qty_total = count_result[1] if count_result else 0
        
        if records_count == 0:
            print(f"[DB] ℹ️ Нет плит для возврата из плана {plan_id}")
            return 0
        
        # Возвращаем все плиты плана в производство
        cur.execute('''
            UPDATE kp_plates
            SET status = 'в производстве', plan_id = NULL
            WHERE plan_id = ?
        ''', (plan_id,))
        
        conn.commit()
        print(f"[DB] ✅ Возвращено {records_count} записей ({qty_total} плит) из плана {plan_id} в производство")
        return records_count
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при возврате плит плана в производство: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def recover_stuck_plates(db_path: str = DEFAULT_DB) -> int:
    """
    Возвращает "застрявшие" плиты обратно в производство.
    
    Простыми словами:
    - Застрявшая плита = статус 'в плане', но не была списана
    - Эти плиты не попали в tracks при планировании
    - Возвращаем их в 'в производстве', чтобы можно было заново запланировать
    
    Когда использовать:
    - После выполнения планов, если часть плит "потерялась"
    - Через команду бота /recover_plates
    
    Аргументы:
        db_path: путь к базе данных
    
    Возвращает:
        Количество восстановленных записей плит
    """
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        # Сначала посмотрим, сколько плит застряло
        cur.execute('''
            SELECT COUNT(*), COALESCE(SUM(qty), 0)
            FROM kp_plates
            WHERE status = 'в плане'
        ''')
        
        count_result = cur.fetchone()
        records_count = count_result[0] if count_result else 0
        qty_total = count_result[1] if count_result else 0
        
        if records_count == 0:
            print("[DB] ℹ️ Нет застрявших плит для восстановления")
            return 0
        
        # Возвращаем все плиты со статусом 'в плане' в производство
        cur.execute('''
            UPDATE kp_plates
            SET status = 'в производстве', plan_id = NULL
            WHERE status = 'в плане'
        ''')
        
        conn.commit()
        print(f"[DB] ✅ Восстановлено {records_count} записей ({qty_total} плит) из статуса 'в плане' в 'в производстве'")
        return records_count
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при восстановлении застрявших плит: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def return_lost_plates_to_production(
    plan_id: str,
    lost_plates: list,
    db_path: str = DEFAULT_DB
) -> int:
    """
    Возвращает конкретные "потерянные" плиты обратно в производство.
    
    ИСПРАВЛЕНО: Теперь обрабатывает ВСЕ записи с одинаковым plate_name,
    а не только первую (убран LIMIT 1).
    
    Простыми словами:
    - Принимает список плит, которые не попали в tracks
    - Возвращает их статус на 'в производстве'
    - Используется как защита при сохранении плана
    - Обрабатывает нужное количество (qty_lost) по нескольким записям
    
    Аргументы:
        plan_id: ID плана (для логирования)
        lost_plates: список словарей [{'kp_id': X, 'plate_name': Y, 'qty_lost': Z}, ...]
        db_path: путь к базе данных
    
    Возвращает:
        Количество восстановленных записей
    """
    if not lost_plates:
        return 0
    
    conn = _connect(db_path)
    total_returned = 0
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        for lost in lost_plates:
            kp_id = lost.get('kp_id')
            plate_name = lost.get('plate_name')
            qty_lost = lost.get('qty_lost', 0)
            
            if not kp_id or not plate_name or qty_lost <= 0:
                continue
            
            # ИСПРАВЛЕНО: Ищем ВСЕ плиты со статусом 'в плане' и данным plan_id (без LIMIT 1)
            cur.execute('''
                SELECT id, qty
                FROM kp_plates
                WHERE kp_id = ? AND plate_name = ? AND status = 'в плане' AND plan_id = ?
                ORDER BY id
            ''', (kp_id, plate_name, plan_id))
            
            rows = cur.fetchall()
            if not rows:
                print(f"[DB] ⚠️ Потерянная плита не найдена: КП #{kp_id}, {plate_name}")
                continue
            
            # Обрабатываем записи по очереди
            remaining_to_return = qty_lost
            
            for row in rows:
                if remaining_to_return <= 0:
                    break
                
                plate_id, current_qty = row
                
                if remaining_to_return >= current_qty:
                    # Возвращаем всю запись в производство
                    cur.execute('''
                        UPDATE kp_plates
                        SET status = 'в производстве', plan_id = NULL
                        WHERE id = ?
                    ''', (plate_id,))
                    print(f"[DB] ✅ Возвращена вся запись: {plate_name} x{current_qty} (запись #{plate_id})")
                    remaining_to_return -= current_qty
                    total_returned += 1
                else:
                    # Частичный возврат: уменьшаем qty в плане, создаём новую запись в производстве
                    new_qty_in_plan = current_qty - remaining_to_return
                    
                    # Обновляем существующую запись
                    cur.execute('''
                        UPDATE kp_plates
                        SET qty = ?
                        WHERE id = ?
                    ''', (new_qty_in_plan, plate_id))
                    
                    insert_kp_plate_remainder_row(
                        cur,
                        source_plate_id=plate_id,
                        remainder_qty=remaining_to_return,
                        status="в производстве",
                        plan_id=None,
                        day_number=None,
                    )

                    print(f"[DB] ✅ Частичный возврат: {plate_name} x{remaining_to_return} (осталось в плане: {new_qty_in_plan}, запись #{plate_id})")
                    remaining_to_return = 0
                    total_returned += 1
        
        conn.commit()
        print(f"[DB] ✅ Всего возвращено {total_returned} записей потерянных плит")
        return total_returned
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при возврате потерянных плит: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


