"""Основная логика планирования производства - оптимизация и распределение по дням"""
import asyncio
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

from core import kp_db
from core.reinforcement_db import get_reinforcement
from core.optimization import optimize_with_cascading_longitudinal_cuts
import core.config_and_data as cfg
import core.optimization as optimization

from ..keyboards import main_menu_kb, calendar_days_kb
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import (
    get_global_day_occupancy,
    MAX_TRACKS_PER_DAY
)

router = Router()


async def load_and_plan_production(message: Message, state: FSMContext):
    """
    Универсальная функция загрузки КП и планирования производства.
    Работает с разными способами фильтрации: date, kp, all, customer.
    """
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    filter_method = data.get('filter_method', 'date')
    
    # === ЗАГРУЗКА КП В ЗАВИСИМОСТИ ОТ ФИЛЬТРА ===
    db_path = PROJECT_ROOT / "plita.db"
    pb_db_path = BOT_DIR / "pb.db"
    
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
    
    # === ДАЛЬШЕ ВСЯ ТЕКУЩАЯ ЛОГИКА ===
    try:
        # === ШАГ 2: СОБИРАЕМ ПЛИТЫ ===
        plates_by_date_and_reinforcement = defaultdict(lambda: defaultdict(list))
        
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
            kp_date = kp_info['date']
            
            cur.execute("""
                SELECT plate_name, length_m, width_m, load_class, qty
                FROM kp_plates
                WHERE kp_id = ? AND status = 'в производстве'
            """, (kp_id,))
            
            for row in cur.fetchall():
                plate_name, length_m, width_m, load_class, qty = row
                load_code = cfg.normalize_load_code(load_class // 100)
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
                    'width': round(width_m * 1000),  # round вместо int для корректного округления
                    'load_code': load_code,
                    'qty': qty,
                    'reinforcement': reinforcement_value,
                    'kp_id': kp_id,
                    'kp_date': kp_date.strftime('%d.%m.%Y'),
                    'customer': kp_info['customer']
                })
        
        # Создаём lookup-таблицы
        plate_to_kp_info = {}
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
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
            
            matching_rests = kp_db.find_matching_rests(
                length_m=length_m,
                width_mm=width_mm,
                qty_needed=qty_needed,
                db_path=plita_db_path
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
                'kp_id': plate_data.get('kp_id')
            })
        
        # ✅ НОВОЕ: Логируем все плиты ДО оптимизации
        logger.info(f"[TRACE] ===== ШАГ 1: ПЛИТЫ ДО ОПТИМИЗАЦИИ =====")
        logger.info(f"[TRACE] Всего плит: {sum(p['qty'] for p in orders_2d)}")
        for p in orders_2d:
            logger.info(f"[TRACE]   {p['plate_name']} × {p['qty']} (длина={p['length']:.2f}м, ширина={p['width']}мм, КП #{p.get('kp_id', '?')})")
        
        # Lookup-таблицы
        plate_lookup_exact = {}
        plate_lookup_by_length = {}
        
        for order in orders_2d:
            key = (round(order['length'], 2), order['width'])
            entry = {
                'kp_date': order.get('kp_date', 'неизвестно'),
                'customer': order.get('customer', 'неизвестно'),
                'plate_name': order.get('plate_name', ''),
                'reinforcement': order.get('reinforcement', 0),
                'load_code': cfg.normalize_load_code(order.get('load_code', 8)),
                'qty_remaining': order.get('qty', 1),
                'kp_id': order.get('kp_id'),
            }
            
            if key not in plate_lookup_exact:
                plate_lookup_exact[key] = []
            plate_lookup_exact[key].append(entry)
            
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
        optimization_result = await asyncio.to_thread(
            optimize_with_cascading_longitudinal_cuts,
            orders_2d=orders_2d
        )
        
        if not optimization_result or optimization_result.get('total_plates', 0) == 0:
            await message.answer(
                "❌ Оптимизация не дала результатов.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
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
        
        # === ПРОВЕРКА: вход vs выход оптимизатора ===
        input_plates = sum(p['qty'] for p in orders_2d)
        output_plates = len(optimization_result.get('plate_assignments', []))
        if input_plates != output_plates:
            logger.error(
                f"[OPT_CHECK] ❌ РАСХОЖДЕНИЕ! Запрошено плит: {input_plates}, "
                f"получено из оптимизатора: {output_plates}, потеряно: {input_plates - output_plates}"
            )
            ordered = Counter(
                (round(o['length'], 2), o['width'], cfg.normalize_load_code(o.get('load_code', 8)))
                for o in orders_2d for _ in range(o['qty'])
            )
            def _plate_key(p):
                length = round(p.get('length', 0), 2)
                width = p.get('width', 0)
                load = cfg.normalize_load_code(p.get('load_code', 8))
                return (length, width, load)
            produced = Counter(_plate_key(p) for p in optimization_result.get('plate_assignments', []))
            missing = ordered - produced
            if missing:
                logger.error(f"[OPT_CHECK] Не хватает в результате: {dict(missing)}")
        else:
            logger.info(f"[OPT_CHECK] ✅ Совпадение: запрошено {input_plates} плит, получено {output_plates}")
        
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
        optimization.OPT_CASCADING_PLAN = optimization_result
        
        all_loads = set(p['load_code'] for p in orders_2d)
        optimization_result['loads_in_group'] = sorted(all_loads)
        optimization.OPT_CASCADING_PLAN_BY_LOAD = {'all': optimization_result}
        optimization.LOAD_TO_REINFORCEMENT_MAP = {
            load_code: ['all'] for load_code in all_loads
        }
        
        cfg.PLATES_1_2 = []
        cfg.PLATE_LOAD_DETAILS = {}
        
        for plate_data in orders_2d:
            length = plate_data['length']
            width_m = plate_data['width'] / 1000.0
            load_code = cfg.normalize_load_code(plate_data['load_code'])
            
            key = (length, width_m, load_code)
            if key in cfg.PLATE_LOAD_DETAILS:
                cfg.PLATE_LOAD_DETAILS[key] += plate_data['qty']
            else:
                cfg.PLATE_LOAD_DETAILS[key] = plate_data['qty']
            
            if abs(width_m - 1.2) < 0.01:
                for _ in range(plate_data['qty']):
                    cfg.PLATES_1_2.append(length)
        
        # === ШАГ 6: ПОДСЧЕТ ДОРОЖЕК ===
        await message.answer("⏳ Подсчитываю дорожки...")
        
        from viz_modules.layout_sequence import build_layout_sequence
        from core.visualization import split_sequence_into_tracks
        
        seq = build_layout_sequence()
        all_tracks_list = split_sequence_into_tracks(seq)

        # === ШАГ 6.5: ЗАЩИТА ОТ ПОТЕРИ ПЛИТ (РЕСКЬЮ) ===
        def _count_tracks_for_rescue(tracks_list):
            counts = {}
            for track in tracks_list:
                for item in track.get('items', []):
                    if not item:
                        continue
                    length = round(item.get('length', 0), 2)
                    load_code = cfg.normalize_load_code(item.get('load_code', 8))
                    mode = item.get('mode', 'solid')
                    if mode == 'split':
                        width_mm = round(item.get('main_w', 1.2) * 1000)
                    elif mode == 'transverse':
                        width_mm = round(item.get('width', 1.2) * 1000)
                    else:
                        width_mm = round(item.get('width', 1.2) * 1000)
                    key = (length, width_mm, load_code)
                    counts[key] = counts.get(key, 0) + 1
                    for sec_cut in item.get('secondary_cuts', []) or []:
                        sec_width_mm = round(sec_cut.get('width', 0) * 1000)
                        sec_length = sec_cut.get('target_length') or length
                        if sec_width_mm > 0:
                            sec_key = (round(sec_length, 2), sec_width_mm, load_code)
                            counts[sec_key] = counts.get(sec_key, 0) + 1
            return counts

        def _build_order_info_map(orders_list):
            info_map = {}
            for order in orders_list:
                key = (
                    round(order.get('length', 0), 2),
                    order.get('width', 1200),
                    cfg.normalize_load_code(order.get('load_code', 8))
                )
                if key not in info_map:
                    info_map[key] = []
                info_map[key].append({
                    'kp_id': order.get('kp_id'),
                    'customer': order.get('customer'),
                    'kp_date': order.get('kp_date'),
                    'plate_name': order.get('plate_name', ''),
                    'qty_remaining': order.get('qty', 1)
                })
            return info_map

        def _create_rescue_tracks(missing_counts, info_map):
            rescue_tracks = []
            current_track = []
            current_len = 0.0
            max_len = 101.0

            def _flush_track():
                nonlocal current_track, current_len
                if current_track:
                    rescue_tracks.append({
                        'items': current_track,
                        'length': current_len,
                        'load_code': 0,
                        'label': 'РЕСКЬЮ',
                        'max_reinforcement': 0.0
                    })
                current_track = []
                current_len = 0.0

            for key, qty_missing in missing_counts.items():
                length, width_mm, load_code = key
                width_m = width_mm / 1000.0
                for _ in range(qty_missing):
                    if current_track and current_len + length > max_len:
                        _flush_track()
                    # Достаём информацию о КП (если есть)
                    order_info = None
                    for entry in info_map.get(key, []):
                        if entry.get('qty_remaining', 0) > 0:
                            entry['qty_remaining'] -= 1
                            order_info = entry
                            break
                    plate_name = ''
                    if order_info and order_info.get('plate_name'):
                        plate_name = order_info['plate_name']
                    else:
                        plate_name = cfg.make_plate_name(length, width_m, load_code=load_code)
                    current_track.append({
                        'length': length,
                        'mode': 'solid',
                        'width': width_m,
                        'load_code': load_code,
                        'label': plate_name,
                        'reinforcement': 0,
                        'kp_id': order_info.get('kp_id') if order_info else None,
                        'customer': order_info.get('customer') if order_info else None,
                        'kp_date': order_info.get('kp_date') if order_info else None,
                        'plate_name': plate_name
                    })
                    current_len += length
            _flush_track()
            return rescue_tracks

        # Считаем потери
        order_counts = {}
        for order in orders_2d:
            key = (
                round(order.get('length', 0), 2),
                order.get('width', 1200),
                cfg.normalize_load_code(order.get('load_code', 8))
            )
            order_counts[key] = order_counts.get(key, 0) + order.get('qty', 1)

        track_counts = _count_tracks_for_rescue(all_tracks_list)
        missing_counts = {}
        for key, qty_need in order_counts.items():
            qty_have = track_counts.get(key, 0)
            if qty_have < qty_need:
                missing_counts[key] = qty_need - qty_have

        if missing_counts:
            info_map = _build_order_info_map(orders_2d)
            rescue_tracks = _create_rescue_tracks(missing_counts, info_map)
            all_tracks_list.extend(rescue_tracks)
            logger.warning(
                f"[RESCUE] Добавлены дополнительные дорожки: {len(rescue_tracks)}. "
                f"Потеряно плит: {sum(missing_counts.values())}"
            )
        
        total_tracks_count = len(all_tracks_list)
        total_days = math.ceil(total_tracks_count / tracks_count)
        
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
            f"2️⃣ Нажмите «💾 Сохранить план» когда всё готово\n\n"
            f"⚠️ ВАЖНО: План сохраняется только после нажатия кнопки!\n"
            f"Без сохранения он останется только в памяти."
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
            day_date = start_dt + timedelta(days=day_num - 1)
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
