"""Основная логика планирования производства - оптимизация и распределение по дням"""
import asyncio
import json
import logging
import sqlite3
import math
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict, Counter

logger = logging.getLogger(__name__)

from aiogram import Router
from aiogram.types import Message
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from bot.services import kp_persistence as kp_db
from core.db_config import PB_DB_PATH, PLITA_DB_PATH
from core.reinforcement_db import get_reinforcement
from core.concrete_grade_resolver import enrich_orders_2d_concrete_grade
from core.work_calendar import nth_working_day
import core.config_and_data as cfg
from core.config_and_data import PlateOrder, canonical_plate_key
import core.optimization as optimization
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.services.optimization_service import OptimizationService
from app.services.production_planning_service import ProductionPlanningService
from core.optimization.result_contract import is_optimization_success
from core.plate_order_context import PlateOrderContext, run_in_order_context

from ..keyboards import main_menu_kb, calendar_days_kb
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import (
    get_global_day_occupancy,
    MAX_TRACKS_PER_DAY
)

# Phase 5 (P8): RESCUE_TOL_LEN_M / RESCUE_TOL_W_MM / RESCUE_EXHAUSTED_RANK
# удалены — fuzzy-матч больше не используется. Identity берётся из
# plate_assignments, см. core.rescue_tracks.

router = Router()
optimization_service = OptimizationService()
production_planning_service = ProductionPlanningService()


async def load_and_plan_production(
    message: Message,
    state: FSMContext,
    plate_order_ctx: PlateOrderContext,
):
    """
    Универсальная функция загрузки КП и планирования производства.
    Работает с разными способами фильтрации: date, kp, all, customer.
    """
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    filter_method = data.get('filter_method', 'date')
    
    # === ЗАГРУЗКА КП В ЗАВИСИМОСТИ ОТ ФИЛЬТРА ===
    db_path = PLITA_DB_PATH
    pb_db_path = PB_DB_PATH
    kp_db.ensure_schema(str(db_path))
    
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    kp_list = []
    
    if filter_method == 'date':
        # Логика по дате
        target_date_str = data.get('target_date')
        target_date = datetime.fromisoformat(target_date_str)
        date_description = data.get('date_description', target_date.strftime('%d.%m.%Y'))
        
        cur.execute("""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE meta.status = 'в работе'
        """)
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            if not exec_terms:
                continue
            
            try:
                exec_date = datetime.strptime(exec_terms, '%d.%m.%Y')
                if exec_date <= target_date:
                    kp_list.append({
                        'kp_id': kp_id,
                        'date': exec_date,
                        'day': exec_date.day,
                        'customer': customer
                    })
            except:
                continue
    
    elif filter_method == 'kp':
        # Логика по номерам КП
        kp_ids = data.get('kp_ids', [])
        
        placeholders = ','.join('?' * len(kp_ids))
        cur.execute(f"""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE kp.kp_id IN ({placeholders})
              AND meta.status = 'в работе'
        """, kp_ids)
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            exec_date = datetime.strptime(exec_terms, '%d.%m.%Y') if exec_terms else datetime.now()
            kp_list.append({
                'kp_id': kp_id,
                'date': exec_date,
                'day': exec_date.day,
                'customer': customer
            })
    
    elif filter_method == 'all':
        # Логика "все КП в работе"
        cur.execute("""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE meta.status = 'в работе'
        """)
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            exec_date = datetime.strptime(exec_terms, '%d.%m.%Y') if exec_terms else datetime.now()
            kp_list.append({
                'kp_id': kp_id,
                'date': exec_date,
                'day': exec_date.day,
                'customer': customer
            })
    
    elif filter_method == 'customer':
        # Логика по заказчику
        customer_name = data.get('customer_name', '')
        
        cur.execute("""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE meta.status = 'в работе'
              AND kp.customer_name = ?
        """, (customer_name,))
        
        for row in cur.fetchall():
            kp_id, exec_terms, customer = row
            exec_date = datetime.strptime(exec_terms, '%d.%m.%Y') if exec_terms else datetime.now()
            kp_list.append({
                'kp_id': kp_id,
                'date': exec_date,
                'day': exec_date.day,
                'customer': customer
            })
    
    if not kp_list:
        conn.close()
        await message.answer(
            "❌ Нет подходящих КП для производства.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
        return
    
    kp_list.sort(key=lambda x: x['date'])
    
    await message.answer(f"✅ Найдено КП: {len(kp_list)}\nЗагружаю плиты...")
    
    # Опциональный фильтр по id плит (при выборе «По КП» с выбором плит)
    kp_plate_ids_raw = data.get('kp_plate_ids') or {}
    kp_plate_ids = {}
    if isinstance(kp_plate_ids_raw, dict):
        for k, v in kp_plate_ids_raw.items():
            sk = str(k)
            if v is not None and not isinstance(v, list):
                v = list(v) if hasattr(v, '__iter__') and not isinstance(v, str) else []
            kp_plate_ids[sk] = v
    # === ДАЛЬШЕ ВСЯ ТЕКУЩАЯ ЛОГИКА ===
    try:
        # === ШАГ 2: СОБИРАЕМ ПЛИТЫ ===
        plates_by_date_and_reinforcement = defaultdict(lambda: defaultdict(list))
        
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
            kp_date = kp_info['date']
            plate_ids_for_kp = kp_plate_ids.get(str(kp_id)) if kp_plate_ids else None
            if plate_ids_for_kp is not None and len(plate_ids_for_kp) == 0:
                continue
            if plate_ids_for_kp and len(plate_ids_for_kp) > 0:
                placeholders = ','.join('?' * len(plate_ids_for_kp))
                cur.execute(f"""
                    SELECT plate_name, length_m, width_m, load_class, qty, length_dm_raw,
                           COALESCE(concrete_grade, '') AS concrete_grade
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве' AND id IN ({placeholders})
                    ORDER BY position_number, id
                """, (kp_id,) + tuple(plate_ids_for_kp))
            else:
                cur.execute("""
                    SELECT plate_name, length_m, width_m, load_class, qty, length_dm_raw,
                           COALESCE(concrete_grade, '') AS concrete_grade
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве'
                """, (kp_id,))
            
            for row in cur.fetchall():
                # length_dm_raw может отсутствовать в старых БД — берём по индексу
                plate_name, length_m, width_m, load_class, qty = row[0], row[1], row[2], row[3], row[4]
                length_dm_raw = (row[5] or '') if len(row) > 5 else ''
                cg_row = str(row[6] or '').strip() if len(row) > 6 else ''
                load_code = cfg.normalize_load_code(load_class // 100)
                # Коррекция ширины по названию: "-12-8п" / "-12-" = 12 дм = 1200 мм; если в БД ошибочно 0.665 — подставляем 1200
                width_mm = round((width_m or 0) * 1000)
                if (plate_name and ("-12-8п" in plate_name or "-12-" in plate_name)) and (width_m is None or width_m < 0.9 or width_m > 1.5):
                    width_mm = 1200
                reinforcement_value = get_reinforcement(
                    length_m=length_m,
                    load_code=load_code,
                    source='series',
                    db_path=pb_db_path,
                    allow_fallback=True
                )
                
                if reinforcement_value is None:
                    reinforcement_value = 999.0
                
                plates_by_date_and_reinforcement[kp_date][reinforcement_value].append({
                    'plate_name': plate_name,
                    'length': length_m,
                    'width': width_mm,  # с учётом коррекции по названию для -12-8п
                    'load_code': load_code,
                    'qty': qty,
                    'reinforcement': reinforcement_value,
                    'kp_id': kp_id,
                    'kp_date': kp_date.strftime('%d.%m.%Y'),
                    'customer': kp_info['customer'],
                    'length_dm_raw': length_dm_raw,
                    'concrete_grade': cg_row or None,
                })
        
        # Создаём lookup-таблицы
        plate_to_kp_info = {}
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
            plate_ids_for_kp = kp_plate_ids.get(str(kp_id)) if kp_plate_ids else None
            if plate_ids_for_kp is not None and len(plate_ids_for_kp) == 0:
                continue
            if plate_ids_for_kp and len(plate_ids_for_kp) > 0:
                placeholders = ','.join('?' * len(plate_ids_for_kp))
                cur.execute(f"""
                    SELECT plate_name, length_m, width_m
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве' AND id IN ({placeholders})
                """, (kp_id,) + tuple(plate_ids_for_kp))
            else:
                cur.execute("""
                    SELECT plate_name, length_m, width_m
                    FROM kp_plates
                    WHERE kp_id = ? AND status = 'в производстве'
                """, (kp_id,))
            for row in cur.fetchall():
                plate_name, length_m, width_m = row
                key = (round(length_m, 2), round(width_m * 1000))  # round для корректного округления
                if key not in plate_to_kp_info:
                    plate_to_kp_info[key] = {
                        'kp_id': kp_id,
                        'kp_date': kp_info['date'].strftime('%d.%m.%Y'),
                        'customer': kp_info['customer'],
                        'plate_name': plate_name
                    }
        
        conn.close()
        
        # === ШАГ 3: ФОРМИРУЕМ СПИСОК ПЛИТ ===
        selected_plates = []
        sorted_dates = sorted(plates_by_date_and_reinforcement.keys())
        
        for kp_date in sorted_dates:
            reinforcement_dict = plates_by_date_and_reinforcement[kp_date]
            sorted_reinforcements = sorted(reinforcement_dict.keys())
            
            for reinforcement in sorted_reinforcements:
                plates = reinforcement_dict[reinforcement]
                for plate_data in plates:
                    selected_plates.append(plate_data)
        
        if not selected_plates:
            await message.answer(
                "❌ Не найдено плит для планирования.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        # === ШАГ 3.5: ПРОВЕРКА ОСТАТКОВ НА СКЛАДЕ ===
        plates_from_rests = []
        plates_for_optimizer = []
        
        plita_db_path = str(PROJECT_ROOT / 'plita.db')
        for plate_data in selected_plates:
            length_m = plate_data['length']
            width_mm = plate_data['width']
            qty_needed = plate_data['qty']
            
            matching_rests = production_planning_service.find_matching_rests(
                length_m=length_m,
                width_mm=width_mm,
                qty_needed=qty_needed,
                db_path=plita_db_path,
            )
            
            qty_from_rests = 0
            
            if matching_rests:
                for rest_info in matching_rests:
                    qty_to_use = rest_info['qty_to_use']
                    qty_from_rests += qty_to_use
                    
                    plates_from_rests.append({
                        'plate_name': plate_data.get('plate_name', ''),
                        'length_m': length_m,
                        'width_mm': width_mm,
                        'qty': qty_to_use,
                        'kp_id': plate_data.get('kp_id'),
                        'kp_date': plate_data.get('kp_date', 'неизвестно'),
                        'customer': plate_data.get('customer', 'неизвестно'),
                        'load_code': cfg.normalize_load_code(plate_data.get('load_code', 8)),
                        'reinforcement': plate_data.get('reinforcement', 0),
                        'rest_id': rest_info['rest_id'],
                        'rest_length': rest_info['rest_length'],
                        'rest_width_mm': rest_info['rest_width_mm'],
                        'match_type': rest_info['match_type'],
                        'cut_cost': rest_info['cut_cost'],
                        'source_plate_name': rest_info['source_plate_name'],
                        'source_kp_id': rest_info['source_kp_id'],
                        'from_rest': True
                    })
            
            qty_remaining = qty_needed - qty_from_rests
            if qty_remaining > 0:
                plate_for_opt = plate_data.copy()
                plate_for_opt['qty'] = qty_remaining
                plates_for_optimizer.append(plate_for_opt)
        
        if plates_from_rests:
            total_from_rests = sum(p['qty'] for p in plates_from_rests)
            rests_msg = f"📦 Плиты из остатков: {total_from_rests} шт\n\n"
            
            for p in plates_from_rests:
                rests_msg += f"✅ {p['plate_name']} × {p['qty']} (КП #{p['kp_id']})\n"
                rests_msg += f"   Из остатка: {p['rest_length']}м × {p['rest_width_mm']}мм\n"
                
                if p['match_type'] == 'exact':
                    rests_msg += f"   Точное совпадение, себестоимость: 0 руб.\n"
                elif p['match_type'] == 'width_cut':
                    rests_msg += f"   Резы: продольный ({p['length_m']}м)\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                elif p['match_type'] == 'length_cut':
                    rests_msg += f"   Резы: поперечный\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                else:
                    rests_msg += f"   Резы: продольный ({p['length_m']}м), поперечный\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                
                rests_msg += "\n"
            
            rests_msg += "💰 Эти плиты уже оплачены - чистая прибыль!"
            await message.answer(rests_msg)
        
        if not plates_for_optimizer:
            await message.answer(
                "✅ Все плиты можно взять из остатков!\n"
                "Оптимизация не требуется.",
                reply_markup=main_menu_kb()
            )
            await state.update_data(
                plates_from_rests=plates_from_rests,
                all_from_rests=True
            )
            await state.clear()
            return
        
        selected_plates = plates_for_optimizer
        await message.answer("⏳ Запускаю оптимизацию раскроя...")
        
        # === ШАГ 4: ОПТИМИЗАЦИЯ ===
        orders_2d = []
        for plate_data in selected_plates:
            orders_2d.append({
                'length': plate_data['length'],
                'width': plate_data['width'],
                'qty': plate_data['qty'],
                'load_code': cfg.normalize_load_code(plate_data['load_code']),
                'reinforcement': plate_data['reinforcement'],
                'kp_date': plate_data.get('kp_date', 'неизвестно'),
                'customer': plate_data.get('customer', 'неизвестно'),
                'plate_name': plate_data.get('plate_name', ''),
                'kp_id': plate_data.get('kp_id'),
                'length_dm_raw': plate_data.get('length_dm_raw', '') or '',
                'concrete_grade': plate_data.get('concrete_grade'),
            })
        enrich_orders_2d_concrete_grade(orders_2d, db_path=PB_DB_PATH)


        # Логируем уникальные load_code в orders_2d (план: этап 2.3)
        unique_loads = set(o['load_code'] for o in orders_2d)
        logger.info(f"[DEMAND] Уникальные load_code в orders_2d: {sorted(unique_loads)}")
        if 12 in unique_loads and 12.5 not in unique_loads:
            kp_ids_in_production = list({p.get('kp_id') for p in selected_plates if p.get('kp_id')})
            if kp_ids_in_production:
                try:
                    with kp_db._connect(kp_db.DEFAULT_DB) as conn:
                        cur = conn.cursor()
                        placeholders = ','.join('?' * len(kp_ids_in_production))
                        cur.execute(
                            f"SELECT 1 FROM kp_plates WHERE kp_id IN ({placeholders}) AND load_class = 1250 LIMIT 1",
                            kp_ids_in_production
                        )
                        if cur.fetchone():
                            logger.warning("[DEMAND] ⚠️ В БД есть плиты 12.5п, но в orders_2d только 12п!")
                except Exception:
                    pass

        # ✅ НОВОЕ: Логируем все плиты ДО оптимизации
        logger.info(f"[TRACE] ===== ШАГ 1: ПЛИТЫ ДО ОПТИМИЗАЦИИ =====")
        logger.info(f"[TRACE] Всего плит: {sum(p['qty'] for p in orders_2d)}")
        for p in orders_2d:
            logger.info(f"[TRACE]   {p['plate_name']} × {p['qty']} (длина={p['length']:.2f}м, ширина={p['width']}мм, КП #{p.get('kp_id', '?')})")
        
        # Lookup-таблицы: ключ строго (length, width), без слияния длин (5.7 и 5.71 — разные ключи).
        plate_lookup_exact = {}
        plate_lookup_by_length = {}
        for order in orders_2d:
            L = round(order['length'], 2)
            W = order['width']
            entry = {
                'kp_date': order.get('kp_date', 'неизвестно'),
                'customer': order.get('customer', 'неизвестно'),
                'plate_name': order.get('plate_name', ''),
                'reinforcement': order.get('reinforcement', 0),
                'load_code': cfg.normalize_load_code(order.get('load_code', 8)),
                'qty_remaining': order.get('qty', 1),
                'kp_id': order.get('kp_id'),
            }
            key = (L, W)
            if key not in plate_lookup_exact:
                plate_lookup_exact[key] = []
            plate_lookup_exact[key].append(entry)
        
        for order in orders_2d:
            length_key = round(order['length'], 2)
            length_entry = {
                'kp_date': order.get('kp_date', 'неизвестно'),
                'customer': order.get('customer', 'неизвестно'),
                'plate_name': order.get('plate_name', ''),
                'reinforcement': order.get('reinforcement', 0),
                'qty_remaining': order.get('qty', 1),
                'kp_id': order.get('kp_id'),
            }
            if length_key not in plate_lookup_by_length:
                plate_lookup_by_length[length_key] = []
            plate_lookup_by_length[length_key].append(length_entry)
        
        # Сортируем списки по дате КП
        def parse_date_for_sort(date_str):
            if not date_str or date_str == 'неизвестно':
                return datetime.max
            try:
                return datetime.strptime(date_str, '%d.%m.%Y')
            except:
                return datetime.max
        
        for key in plate_lookup_exact:
            plate_lookup_exact[key].sort(key=lambda x: parse_date_for_sort(x.get('kp_date', '')))
        
        for key in plate_lookup_by_length:
            plate_lookup_by_length[key].sort(key=lambda x: parse_date_for_sort(x.get('kp_date', '')))
        
        # Запуск оптимизации
        optimization_context = await run_in_order_context(
            plate_order_ctx,
            optimization_service.optimize,
            AppPlateOrder.from_orders_2d(orders_2d),
        )
        optimization_result = optimization_context.optimization_result
        if (
            not is_optimization_success(optimization_result)
            or optimization_result.get("total_plates", 0) == 0
        ):
            await message.answer(
                "❌ Оптимизация не дала результатов.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return

        # P8.1: backfill identity у plate_assignments, чтобы slot_exhausted /
        # secondary_unmapped не блокировали commit_plan_plates.
        from core.plate_attribution import backfill_assignment_identity
        _backfilled_count = backfill_assignment_identity(
            optimization_result.get('plate_assignments', []) or [],
            orders_2d,
        )
        if _backfilled_count:
            logger.info(
                "[BOT-PLAN] Восстановлена identity у %s plate_assignments-записей",
                _backfilled_count,
            )

        await message.answer(f"✅ Оптимизация завершена! Исходных плит: {optimization_result.get('total_plates', 0)}")
        
        # ✅ НОВОЕ: Логируем результаты оптимизации
        logger.info(f"[TRACE] ===== ШАГ 2: РЕЗУЛЬТАТЫ ОПТИМИЗАЦИИ =====")
        logger.info(f"[TRACE] Первичных резов: {len(optimization_result.get('primary_cuts', []))}")
        logger.info(f"[TRACE] Вторичных резов: {len(optimization_result.get('secondary_cuts', []))}")
        
        # Подсчитываем плиты по типам
        primary_plates_count = 0
        for cut in optimization_result.get('primary_cuts', []):
            primary_plates_count += cut.get('qty', 0)
            logger.info(f"[TRACE]   Первичный: {cut['width']}мм + {cut['rest']}мм × {cut['qty']} (длины={cut.get('lengths', [])})")
        
        secondary_plates_count = 0
        for cut in optimization_result.get('secondary_cuts', []):
            secondary_plates_count += cut.get('qty', 0) * cut.get('pieces', 1)
            logger.info(f"[TRACE]   Вторичный: {cut['source']}мм → {cut['cuts']} × {cut['qty']} (тип={cut.get('type', '?')})")
        
        logger.info(f"[TRACE] Плит из первичных резов: {primary_plates_count}")
        logger.info(f"[TRACE] Плит из вторичных резов: {secondary_plates_count}")
        logger.info(f"[TRACE] Всего плит после оптимизации: {primary_plates_count + secondary_plates_count}")
        
        # === ПРОВЕРКА: вход vs выход оптимизатора (только первичные резы) ===
        # plate_assignments содержит первичные + вторичные; сравниваем только с первичными,
        # вторичные — бонусные плиты из остатков, их больше, чем заказано, и это ожидаемо.
        input_plates = sum(p['qty'] for p in orders_2d)
        all_assignments = optimization_result.get('plate_assignments', [])
        output_plates_primary = sum(1 for p in all_assignments if p.get('source') == 'primary')
        output_plates_total = len(all_assignments)
        if input_plates != output_plates_primary:
            logger.error(
                f"[OPT_CHECK] ❌ РАСХОЖДЕНИЕ! Запрошено плит: {input_plates}, "
                f"первичных в результате: {output_plates_primary}, "
                f"всего в plate_assignments (вкл. вторичные): {output_plates_total}, "
                f"потеряно первичных: {input_plates - output_plates_primary}"
            )
            ordered = Counter(
                canonical_plate_key(o['length'], o['width'], o.get('load_code', 8))
                for o in orders_2d for _ in range(o['qty'])
            )
            def _plate_key(p):
                return canonical_plate_key(
                    p.get('length', 0), p.get('width', 0), p.get('load_code', 8)
                )
            produced = Counter(
                _plate_key(p) for p in all_assignments if p.get('source') == 'primary'
            )
            missing = ordered - produced
            if missing:
                logger.error(f"[OPT_CHECK] Не хватает в результате: {dict(missing)}")
        else:
            logger.info(
                f"[OPT_CHECK] ✅ Совпадение: запрошено {input_plates} плит, "
                f"первичных получено {output_plates_primary}, "
                f"вторичных дополнительно: {output_plates_total - output_plates_primary}"
            )
        

        # === ШАГ 4.5: ДОПОЛНЕНИЕ LOOKUP ДЛЯ ВТОРИЧНЫХ РЕЗОВ ===
        if optimization_result.get('secondary_cuts'):
            orders_dict = {}
            for order in orders_2d:
                key = (
                    round(order['length'], 2),
                    order['width'],
                    cfg.normalize_load_code(order.get('load_code', 8))
                )
                if key not in orders_dict:
                    orders_dict[key] = []
                orders_dict[key].append(order)
            
            for sec_cut in optimization_result['secondary_cuts']:
                target_key = sec_cut.get('target_order_key')
                if not target_key:
                    continue
                
                target_length, target_width, target_load_code = target_key
                original_orders = orders_dict.get((round(target_length, 2), target_width, target_load_code), [])
                
                if not original_orders:
                    continue
                
                result_lengths = sec_cut.get('lengths', [])
                result_width = sec_cut['cuts'][0]
                
                for result_length in result_lengths:
                    key_result = (round(result_length, 2), result_width)
                    
                    original_order = None
                    for order in original_orders:
                        original_key = (round(order['length'], 2), order['width'])
                        if original_key in plate_lookup_exact:
                            for entry in plate_lookup_exact[original_key]:
                                if entry.get('qty_remaining', 0) > 0:
                                    original_order = order
                                    break
                        if original_order:
                            break
                    
                    if not original_order:
                        continue
                    
                    if key_result not in plate_lookup_exact:
                        plate_lookup_exact[key_result] = []
                    
                    plate_lookup_exact[key_result].append({
                        'kp_date': original_order.get('kp_date', 'неизвестно'),
                        'customer': original_order.get('customer', 'неизвестно'),
                        'plate_name': original_order.get('plate_name', ''),
                        'reinforcement': original_order.get('reinforcement', 0),
                        'load_code': cfg.normalize_load_code(original_order.get('load_code', 8)),
                        'qty_remaining': 1,
                        'kp_id': original_order.get('kp_id'),
                        'is_from_secondary': True
                    })
                    
                    length_key = round(result_length, 2)
                    if length_key not in plate_lookup_by_length:
                        plate_lookup_by_length[length_key] = []
                    
                    plate_lookup_by_length[length_key].append({
                        'kp_date': original_order.get('kp_date', 'неизвестно'),
                        'customer': original_order.get('customer', 'неизвестно'),
                        'plate_name': original_order.get('plate_name', ''),
                        'reinforcement': original_order.get('reinforcement', 0),
                        'qty_remaining': 1,
                        'kp_id': original_order.get('kp_id'),
                        'is_from_secondary': True
                    })
            
            logger.info(
                f"[PRODUCTION] Дополнено lookup-таблиц: exact={len(plate_lookup_exact)}, by_length={len(plate_lookup_by_length)}"
            )
        
        # === ШАГ 5: ПОДГОТОВКА ДЛЯ ВИЗУАЛИЗАЦИИ ===
        all_loads = set(p['load_code'] for p in orders_2d)
        plate_order_ctx.load_production_snapshot(orders_2d, optimization_result)
        optimization_result['loads_in_group'] = sorted(all_loads)
        

        # === ШАГ 6: ПОДСЧЕТ ДОРОЖЕК ===
        await message.answer("⏳ Подсчитываю дорожки...")
        
        from viz_modules.layout_sequence import build_layout_sequence
        from core.visualization import (
            LayoutIntegrityError,
            TrackLayoutInvariantError,
            split_sequence_into_tracks,
        )
        from core.plate_audit import PlateAudit as _PlateAudit

        # Подхватываем audit из оптимизатора, если он там был создан
        _handler_audit: _PlateAudit | None = optimization_result.get('_plate_audit')
        if _handler_audit is None:
            _handler_audit = _PlateAudit(orders_2d)

        seq = build_layout_sequence()

        # PlateAudit: checkpoint после build_layout_sequence
        _handler_audit.checkpoint("layout_sequence", seq)

        try:
            all_tracks_list = split_sequence_into_tracks(
                seq,
                strict_layout_integrity=True,
            )
        except LayoutIntegrityError as exc:
            logger.error("[BOT-PLAN] Ошибка целостности раскладки: %s", exc)
            await message.answer(
                f"❌ Нарушена целостность раскладки дорожек: {exc}"
            )
            return
        except TrackLayoutInvariantError as exc:
            logger.error("[BOT-PLAN] Нарушены правила старта дорожки с целой плиты: %s", exc)
            await message.answer(
                f"❌ Невозможно разложить дорожки без целой плиты в начале: {exc}"
            )
            return

        from core.config.settings import get_settings

        if get_settings().track_top_up_from_following:
            from core.track_top_up import top_up_tracks_from_following

            top_up_tracks_from_following(all_tracks_list or [])

        # PlateAudit: checkpoint после split_sequence_into_tracks
        _handler_audit.checkpoint("tracks", all_tracks_list)


        # === ШАГ 6.5: ЗАЩИТА ОТ ПОТЕРИ ПЛИТ (РЕСКЬЮ) ===
        # Phase 5 (P8): локальная rescue-логика удалена. Источник правды —
        # plate_assignments (с backfill identity, см. core.plate_attribution).
        # build_rescue_tracks принимает plate_assignments, считает дефицит по
        # точной identity (kp_id, plate_name) — fuzzy-матч RESCUE_TOL_* ушёл.
        from core.rescue_tracks import build_rescue_tracks

        plate_assignments = optimization_result.get('plate_assignments', []) or []
        rescue_tracks, missing_counts, rescue_assignments = build_rescue_tracks(
            orders_2d=orders_2d,
            plate_assignments=plate_assignments,
        )
        rescue_tracks_added = 0
        if rescue_tracks:
            all_tracks_list.extend(rescue_tracks)
            rescue_tracks_added = len(rescue_tracks)
        if rescue_assignments:
            optimization_result.setdefault('plate_assignments', []).extend(
                rescue_assignments
            )
        logger.info(
            "[RESCUE] missing identities=%s, rescue_tracks=%s, rescue_assignments=%s",
            len(missing_counts), rescue_tracks_added, len(rescue_assignments),
        )

        # P9: backfill identity у track-items / secondary_cuts. Без него
        # secondary без kp_id выпадают из _count_track_items_by_day, плиты
        # помечаются 'в плане' с day_number=NULL и не списываются.
        from core.plate_attribution import backfill_track_items_identity
        _backfilled_items_count = backfill_track_items_identity(
            all_tracks_list,
            orders_2d,
        )
        if _backfilled_items_count:
            logger.info(
                "[BOT-PLAN] Восстановлена identity у %s track items "
                "(root + secondary_cuts)",
                _backfilled_items_count,
            )
        # PlateAudit: финальный checkpoint после rescue
        _handler_audit.checkpoint("final", all_tracks_list)
        if _handler_audit.has_losses("input", "final"):
            logger.error("[AUDIT] Итоговые потери плит:\n%s", _handler_audit.summary())
        else:
            logger.info("[AUDIT] Все плиты учтены:\n%s", _handler_audit.summary())

        total_tracks_count = len(all_tracks_list)
        total_days = math.ceil(total_tracks_count / tracks_count)
        
        # Проверка: все плиты плана есть в БД (этап 3 — не придумываем плиты).
        # Phase 5: считаем плиты по всем track items (включая secondary_cuts).
        # Раньше вызывался удалённый _raw_track_counts.
        def _count_plan_plates(tracks: list) -> Counter:
            counts: Counter = Counter()
            for track in tracks or []:
                for item in track.get('items') or []:
                    if not item:
                        continue
                    L = round(float(item.get('length') or 0), 2)
                    W_raw = item.get('width') or item.get('main_w') or 1.2
                    if isinstance(W_raw, (int, float)) and 0 < W_raw < 10:
                        W = int(round(float(W_raw) * 1000))
                    else:
                        W = int(round(float(W_raw or 1200)))
                    LC = cfg.normalize_load_code(item.get('load_code', 8))
                    if L > 0 and W > 0:
                        counts[(L, W, LC)] += 1
                    for sec in item.get('secondary_cuts') or []:
                        if not sec:
                            continue
                        sl = round(float(sec.get('target_length') or L or 0), 2)
                        sw_raw = sec.get('width', 0)
                        if isinstance(sw_raw, (int, float)) and 0 < sw_raw < 10:
                            sw = int(round(float(sw_raw) * 1000))
                        else:
                            sw = int(round(float(sw_raw or 0)))
                        slc = cfg.normalize_load_code(sec.get('load_code', LC))
                        if sl > 0 and sw > 0:
                            counts[(sl, sw, slc)] += 1
            return counts

        plan_plates = _count_plan_plates(all_tracks_list)
        kp_ids_in_production = list({p.get('kp_id') for p in selected_plates if p.get('kp_id')})
        if kp_ids_in_production and plan_plates:
            try:
                db_plates = Counter()
                with kp_db._connect(kp_db.DEFAULT_DB) as conn:
                    cur = conn.cursor()
                    placeholders = ','.join('?' * len(kp_ids_in_production))
                    cur.execute(f"""
                        SELECT length_m, width_m, load_class, SUM(qty)
                        FROM kp_plates
                        WHERE kp_id IN ({placeholders})
                        AND status IN ('в производстве', 'в плане')
                        GROUP BY length_m, width_m, load_class
                    """, kp_ids_in_production)
                    for row in cur.fetchall():
                        length_m, width_m, load_class, qty = row
                        lc = cfg.normalize_load_code(load_class // 100)
                        width_mm = int(round(width_m * 1000))
                        db_plates[(round(length_m, 2), width_mm, lc)] += qty
                extra_in_plan = plan_plates - db_plates
                missing_in_plan = db_plates - plan_plates
                if extra_in_plan:
                    logger.error(f"[ПЛАН] В плане есть плиты, которых нет в БД (придуманные): {dict(extra_in_plan)}")
                if missing_in_plan:
                    logger.warning(f"[ПЛАН] В БД есть плиты, не попавшие в план (остатки?): {dict(missing_in_plan)}")
            except Exception as e:
                logger.exception(f"[ПЛАН] Проверка план vs БД: {e}")
        
        # === СОХРАНЯЕМ ВСЕ ДАННЫЕ ===
        target_date_str = data.get('target_date')
        plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
        completed_days = data.get('completed_days', [])
        
        # Форматируем дату начала для читаемости
        plan_start_display = plan_start_date
        try:
            plan_start_display = datetime.strptime(plan_start_date, '%Y-%m-%d').strftime('%d.%m.%Y')
        except:
            pass
        
        await message.answer(
            f"✅ План готов!\n\n"
            f"📊 Параметры:\n"
            f"  • Дата начала: {plan_start_display}\n"
            f"  • Всего дорожек: {total_tracks_count}\n"
            f"  • Дорожек в день: {tracks_count}\n"
            f"  • Потребуется дней: {total_days}\n\n"
            f"💡 Что дальше?\n"
            f"1️⃣ Просмотрите дни ниже 👇\n"
            f"2️⃣ Чтобы посмотреть диаграмму ДО сохранения — нажмите «📈 Диаграмма этого плана»\n"
            f"3️⃣ Нажмите «💾 Сохранить план» когда всё готово\n\n"
            f"⚠️ ВАЖНО: План сохраняется только после нажатия кнопки!\n"
            f"Без сохранения он останется только в памяти.\n\n"
            f"«📈 Диаграмма этого плана» — по текущему расчёту (даже без сохранения)."
        )
        
        # Рассчитываем days_info с глобальной загруженностью для новых дат
        global_occupancy = get_global_day_occupancy()
        
        # Создаём days_info для каждого дня нового плана
        days_info = {}
        try:
            start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
        except:
            start_dt = datetime.now()
        
        # Проверяем, не превышает ли план доступные слоты
        overloaded_days = []  # Дни, где будет превышение
        
        for day_num in range(1, total_days + 1):
            day_date = datetime.combine(
                nth_working_day(start_dt.date(), day_num),
                datetime.min.time(),
            )
            date_key = day_date.strftime('%Y-%m-%d')
            date_display = day_date.strftime('%d.%m')
            
            current_occupied = global_occupancy.get(date_key, 0)
            free_slots = MAX_TRACKS_PER_DAY - current_occupied
            
            # Проверяем превышение
            if tracks_count > free_slots:
                overloaded_days.append({
                    'date': date_display,
                    'occupied': current_occupied,
                    'free': free_slots,
                    'want': tracks_count,
                    'excess': tracks_count - free_slots
                })
            
            days_info[date_key] = {
                'occupied': current_occupied,
                'max': MAX_TRACKS_PER_DAY,
                'completed': False,
                'day_number': day_num
            }
        
        # Если есть превышение - показываем предупреждение
        if overloaded_days:
            warning_lines = ["⚠️ ВНИМАНИЕ! Превышение лимита дорожек!\n"]
            warning_lines.append(f"Вы хотите планировать {tracks_count} дорожек/день,")
            warning_lines.append(f"но на некоторых датах не хватает места:\n")
            
            for day in overloaded_days[:5]:  # Показываем максимум 5 проблемных дней
                warning_lines.append(
                    f"  • {day['date']}: занято {day['occupied']}/5, "
                    f"свободно {day['free']}, нужно {day['want']}"
                )
            
            if len(overloaded_days) > 5:
                warning_lines.append(f"  ... и ещё {len(overloaded_days) - 5} дней")
            
            warning_lines.append(f"\n💡 Решения:")
            warning_lines.append(f"1️⃣ Уменьшите количество дорожек в день")
            warning_lines.append(f"2️⃣ Выберите другую дату начала")
            warning_lines.append(f"3️⃣ Удалите или отредактируйте другие планы")
            
            await message.answer('\n'.join(warning_lines))
        
        await state.update_data(
            total_tracks_count=total_tracks_count,
            total_days=total_days,
            tracks_count=tracks_count,
            all_tracks_list=all_tracks_list,
            orders_2d=orders_2d,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            plate_to_kp_info=plate_to_kp_info,
            optimization_result=optimization_result,
            target_date=target_date_str,
            plates_from_rests=plates_from_rests,
            plan_start_date=plan_start_date,
            completed_days=completed_days,
            days_info=days_info
        )

        
        # === ПОКАЗЫВАЕМ КНОПКИ С ДАТАМИ ===
        await message.answer(
            "Выберите день производства:",
            reply_markup=calendar_days_kb(
                total_days, 
                plan_start_date, 
                completed_days,
                days_info
            )
        )
        
        await state.set_state(ProductionStates.waiting_day_selection)
        
    except Exception as e:
        logger.exception(f"Ошибка при планировании производства: {e}")
        await message.answer(
            "❌ Ошибка при планировании производства.\n\n"
            "Попробуйте позже. Если повторяется — смотри logs/bot.log.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
