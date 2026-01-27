"""Основная логика планирования производства - оптимизация и распределение по дням"""
import asyncio
import logging
import sqlite3
import math
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

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
                load_code = load_class // 100
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
                    'width': int(width_m * 1000),
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
                key = (round(length_m, 2), int(width_m * 1000))
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
                        'load_code': plate_data.get('load_code', 8),
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
                'load_code': plate_data['load_code'],
                'reinforcement': plate_data['reinforcement'],
                'kp_date': plate_data.get('kp_date', 'неизвестно'),
                'customer': plate_data.get('customer', 'неизвестно'),
                'plate_name': plate_data.get('plate_name', ''),
                'kp_id': plate_data.get('kp_id')
            })
        
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
                'load_code': order.get('load_code', 8),
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
        
        # === ШАГ 4.5: ДОПОЛНЕНИЕ LOOKUP ДЛЯ ВТОРИЧНЫХ РЕЗОВ ===
        if optimization_result.get('secondary_cuts'):
            orders_dict = {}
            for order in orders_2d:
                key = (round(order['length'], 2), order['width'])
                if key not in orders_dict:
                    orders_dict[key] = []
                orders_dict[key].append(order)
            
            for sec_cut in optimization_result['secondary_cuts']:
                target_key = sec_cut.get('target_order_key')
                if not target_key:
                    continue
                
                target_length, target_width = target_key
                original_orders = orders_dict.get((round(target_length, 2), target_width), [])
                
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
                        'load_code': original_order.get('load_code', 8),
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
            load_code = plate_data['load_code']
            
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
