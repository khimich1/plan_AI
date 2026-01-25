"""Обработчики планирования производства плит"""
import asyncio
import logging
import os
import json
from datetime import datetime
from pathlib import Path
from collections import defaultdict
import math

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import Message, FSInputFile, CallbackQuery
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
from core.reinforcement_db import get_reinforcement
from core.visualization import visualize_plan
from core.optimization import optimize_with_cascading_longitudinal_cuts
from core.formovka_excel import create_formovka_files_for_tracks
from core.gantt_excel import create_gantt_excel
import core.config_and_data as cfg
import core.optimization as optimization

from ..keyboards import main_menu_kb, production_days_kb, production_day_actions_kb, plates_completion_kb, cancel_process_kb, production_filter_kb
from ..states import ProductionStates
from ..bot_config import OUTPUTS_DIR_STR

router = Router()


@router.message(F.text == "Планирование производства")
async def btn_production_planning(message: Message, state: FSMContext):
    """Обработчик кнопки 'Планирование производства'"""
    await state.set_state(ProductionStates.waiting_tracks_count)
    await message.answer(
        "📋 Планирование производства плит\n\n"
        "Шаг 1 из 2: Сколько дорожек нужно загрузить сегодня?\n"
        "(Введите число, например: 5)",
        reply_markup=main_menu_kb()
    )
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
    )


@router.message(ProductionStates.waiting_tracks_count)
async def receive_tracks_count(message: Message, state: FSMContext):
    """Получаем количество дорожек"""
    try:
        tracks_count = int(message.text.strip())
        
        if tracks_count <= 0 or tracks_count > 50:
            await message.answer(
                "❌ Количество дорожек должно быть от 1 до 50.\n"
                "Попробуйте снова:"
            )
            await message.answer(
                "Или нажмите кнопку ниже для отмены:",
                reply_markup=cancel_process_kb()
            )
            return
    except ValueError:
        await message.answer(
            "❌ Неверный формат. Введите целое число (например: 5):"
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return
    
    await state.update_data(tracks_count=tracks_count)
    
    # ИЗМЕНЕНИЕ: переходим к выбору способа фильтрации
    await state.set_state(ProductionStates.waiting_filter_method)
    await message.answer(
        f"✅ Дорожек: {tracks_count}\n\n"
        "Шаг 2 из 3: Как выбрать плиты для производства?\n\n"
        "Выберите способ:",
        reply_markup=production_filter_kb()
    )


# === ОБРАБОТЧИКИ ВЫБОРА СПОСОБА ФИЛЬТРАЦИИ ===

@router.callback_query(F.data == "filter_by_date", ProductionStates.waiting_filter_method)
async def filter_by_date(callback: CallbackQuery, state: FSMContext):
    """Выбор по дате - показываем текущую логику"""
    await state.update_data(filter_method='date')
    await state.set_state(ProductionStates.waiting_date_number)
    await callback.message.answer(
        "Шаг 3 из 3: До какой даты брать плиты?\n\n"
        "Поддерживаемые форматы:\n"
        "• 25 (число текущего месяца)\n"
        "• 01.02.2026 (полная дата)\n"
        "• 2026-02-01 (ISO формат)\n\n"
        "Введите дату:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "filter_by_kp", ProductionStates.waiting_filter_method)
async def filter_by_kp(callback: CallbackQuery, state: FSMContext):
    """Выбор по номерам КП"""
    await state.update_data(filter_method='kp')
    await state.set_state(ProductionStates.waiting_kp_numbers)
    await callback.message.answer(
        "Шаг 3 из 3: Введите номера КП\n\n"
        "Поддерживаемые форматы:\n"
        "• 15 (один номер)\n"
        "• 15, 18, 22 (несколько через запятую)\n"
        "• 15-22 (диапазон)\n\n"
        "Введите номера:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "filter_all", ProductionStates.waiting_filter_method)
async def filter_all(callback: CallbackQuery, state: FSMContext):
    """Выбор всех КП в работе - сразу переходим к загрузке"""
    await state.update_data(filter_method='all')
    # Сразу вызываем функцию загрузки (без ввода даты)
    await load_and_plan_production(callback.message, state)
    await callback.answer()


@router.callback_query(F.data == "filter_by_customer", ProductionStates.waiting_filter_method)
async def filter_by_customer(callback: CallbackQuery, state: FSMContext):
    """Выбор по заказчику - показываем список заказчиков"""
    await state.update_data(filter_method='customer')
    
    # Загружаем список уникальных заказчиков из БД
    db_path = PROJECT_ROOT / "plita.db"
    import sqlite3
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute('''
        SELECT DISTINCT customer_name 
        FROM KP_offers 
        WHERE customer_name IS NOT NULL
        ORDER BY customer_name
    ''')
    customers = [row[0] for row in cur.fetchall()]
    conn.close()
    
    if not customers:
        await callback.message.answer(
            "❌ Нет заказчиков в базе данных.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
        await callback.answer()
        return
    
    # Создаем клавиатуру с заказчиками
    from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
    buttons = []
    for customer in customers:
        buttons.append([
            InlineKeyboardButton(
                text=customer,
                callback_data=f"customer_{customer}"
            )
        ])
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад", callback_data="cancel_process")
    ])
    
    await callback.message.answer(
        "Выберите заказчика:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )
    await callback.answer()


@router.callback_query(F.data.startswith("customer_"), ProductionStates.waiting_filter_method)
async def receive_customer_selection(callback: CallbackQuery, state: FSMContext):
    """Обработчик выбора заказчика"""
    customer_name = callback.data.replace("customer_", "")
    
    await state.update_data(customer_name=customer_name)
    
    await callback.message.answer(f"✅ Заказчик: {customer_name}\n\n⏳ Загружаю плиты...")
    await load_and_plan_production(callback.message, state)
    await callback.answer()


# === ОБРАБОТЧИКИ ВВОДА НОМЕРОВ КП ===

@router.message(ProductionStates.waiting_kp_numbers)
async def receive_kp_numbers(message: Message, state: FSMContext):
    """Парсим номера КП и запускаем планирование"""
    user_input = message.text.strip()
    
    kp_ids = []
    
    try:
        # Формат 1: Диапазон "15-22"
        if '-' in user_input and user_input.count('-') == 1:
            parts = user_input.split('-')
            start_str = parts[0].strip()
            end_str = parts[1].strip()
            
            # Проверка что это не дата (не содержит точки)
            if '.' not in start_str and '.' not in end_str:
                start_num = int(start_str)
                end_num = int(end_str)
                if start_num > end_num:
                    raise ValueError("Начало диапазона больше конца")
                kp_ids = list(range(start_num, end_num + 1))
            else:
                raise ValueError("Формат не распознан как диапазон номеров КП")
        
        # Формат 2: Список "15, 18, 22"
        elif ',' in user_input:
            kp_ids = [int(x.strip()) for x in user_input.split(',')]
        
        # Формат 3: Одно число "15"
        else:
            kp_ids = [int(user_input)]
        
        if not kp_ids:
            raise ValueError("Не указаны номера КП")
        
    except ValueError as e:
        await message.answer(
            f"❌ Неверный формат: {e}\n\n"
            "Попробуйте снова:\n"
            "• 15 (один номер)\n"
            "• 15, 18, 22 (список)\n"
            "• 15-22 (диапазон)"
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return
    
    # Сохраняем номера КП
    await state.update_data(kp_ids=kp_ids)
    
    await message.answer(f"✅ Выбрано КП: {', '.join(map(str, kp_ids))}\n\n⏳ Загружаю плиты...")
    
    # Запускаем загрузку
    await load_and_plan_production(message, state)


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
    
    import sqlite3
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
    # (Копирую из старой функции начиная со строки 376)
    try:
        # === ШАГ 2: СОБИРАЕМ ПЛИТЫ ===
        plates_by_date_and_reinforcement = defaultdict(lambda: defaultdict(list))
        
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
            kp_date = kp_info['date']
            
            cur.execute("""
                SELECT plate_name, length_m, width_m, load_class, qty
                FROM kp_plates
                WHERE kp_id = ?
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
                WHERE kp_id = ?
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
        seq = build_layout_sequence()
        
        MAX_TRACK_LENGTH = 101.0
        all_tracks_list = []
        
        if isinstance(seq, list) and seq and isinstance(seq[0], dict) and 'load_code' in seq[0]:
            for group in seq:
                load_code = group['load_code']
                items = group['sequence']
                group_label = group.get('label', f'Нагрузка {load_code}п')
                
                current_track = []
                current_track_length = 0.0
                max_reinforcement_in_track = 0.0
                
                for i, item in enumerate(items):
                    item_length = item['length']
                    item_reinforcement = item.get('reinforcement', 0) or 0
                    will_exceed = (current_track_length + item_length > MAX_TRACK_LENGTH and current_track)
                    
                    if will_exceed:
                        all_tracks_list.append({
                            'items': current_track,
                            'length': current_track_length,
                            'load_code': load_code,
                            'label': group_label,
                            'max_reinforcement': max_reinforcement_in_track
                        })
                        current_track = []
                        current_track_length = 0.0
                        max_reinforcement_in_track = 0.0
                    
                    current_track.append(item)
                    current_track_length += item_length
                    max_reinforcement_in_track = max(max_reinforcement_in_track, item_reinforcement)
                
                if current_track:
                    all_tracks_list.append({
                        'items': current_track,
                        'length': current_track_length,
                        'load_code': load_code,
                        'label': group_label,
                        'max_reinforcement': max_reinforcement_in_track
                    })
        else:
            current_track = []
            current_track_length = 0.0
            max_reinforcement_in_track = 0.0
            
            for i, item in enumerate(seq):
                item_length = item['length']
                item_reinforcement = item.get('reinforcement', 0) or 0
                will_exceed = (current_track_length + item_length > MAX_TRACK_LENGTH and current_track)
                
                if will_exceed:
                    all_tracks_list.append({
                        'items': current_track,
                        'length': current_track_length,
                        'max_reinforcement': max_reinforcement_in_track
                    })
                    current_track = []
                    current_track_length = 0.0
                    max_reinforcement_in_track = 0.0
                
                current_track.append(item)
                current_track_length += item_length
                max_reinforcement_in_track = max(max_reinforcement_in_track, item_reinforcement)
            
            if current_track:
                all_tracks_list.append({
                    'items': current_track,
                    'length': current_track_length,
                    'max_reinforcement': max_reinforcement_in_track
                })
        
        total_tracks_count = len(all_tracks_list)
        total_days = math.ceil(total_tracks_count / tracks_count)
        
        await message.answer(
            f"📊 Анализ завершен!\n\n"
            f"• Всего дорожек: {total_tracks_count}\n"
            f"• Дорожек в день: {tracks_count}\n"
            f"• Потребуется дней: {total_days}\n\n"
            f"Выберите день для просмотра:"
        )
        
        # === СОХРАНЯЕМ ВСЕ ДАННЫЕ ===
        target_date_str = data.get('target_date')
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
            plates_from_rests=plates_from_rests
        )
        
        # === ПОКАЗЫВАЕМ КНОПКИ ===
        await message.answer(
            "Выберите день производства:",
            reply_markup=production_days_kb(total_days)
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


@router.message(ProductionStates.waiting_date_number)
async def receive_date_number_and_plan(message: Message, state: FSMContext):
    """Парсим дату и запускаем планирование"""
    user_input = message.text.strip()
    
    # === ПАРСИНГ ДАТЫ: поддерживаем разные форматы ===
    target_date = None
    date_description = ""
    
    # Формат 1: Полная дата "ДД.ММ.ГГГГ"
    try:
        target_date = datetime.strptime(user_input, '%d.%m.%Y')
        date_description = target_date.strftime('%d.%m.%Y')
    except ValueError:
        pass
    
    # Формат 2: Полная дата "ГГГГ-ММ-ДД"
    if not target_date:
        try:
            target_date = datetime.strptime(user_input, '%Y-%m-%d')
            date_description = target_date.strftime('%d.%m.%Y')
        except ValueError:
            pass
    
    # Формат 3: Только число месяца
    if not target_date:
        try:
            date_number = int(user_input)
            
            if date_number < 1 or date_number > 31:
                await message.answer(
                    "❌ Число должно быть от 1 до 31.\n"
                    "Попробуйте снова:"
                )
                await message.answer(
                    "Или нажмите кнопку ниже для отмены:",
                    reply_markup=cancel_process_kb()
                )
                return
            
            now = datetime.now()
            target_date = datetime(now.year, now.month, date_number)
            date_description = f"{date_number} {target_date.strftime('%B %Y')}"
        except ValueError:
            pass
    
    if not target_date:
        await message.answer(
            "❌ Неверный формат даты.\n\n"
            "Поддерживаемые форматы:\n"
            "• 25 (число месяца)\n"
            "• 01.02.2026 (полная дата)\n"
            "• 2026-02-01 (ISO формат)\n\n"
            "Попробуйте снова:"
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return
    
    # Сохраняем дату и описание
    await state.update_data(
        target_date=target_date.isoformat(),
        date_description=date_description
    )
    
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    
    await message.answer(
        f"✅ Параметры планирования:\n"
        f"• Дорожек в день: {tracks_count}\n"
        f"• Плиты со сроком до: {date_description}\n\n"
        f"⏳ Загружаю плиты из базы данных..."
    )
    
    # Запускаем универсальную функцию загрузки
    await load_and_plan_production(message, state)


# === ОБРАБОТЧИКИ ВЫБОРА ДНЯ ===

@router.callback_query(F.data.startswith("production_day_"))
async def process_day_selection(callback: CallbackQuery, state: FSMContext):
    """Обработчик выбора конкретного дня"""
    
    day_number = int(callback.data.split("_")[-1])
    
    await callback.message.answer(
        f"📄 Генерирую план для Дня {day_number}...\n"
        f"⏳ Пожалуйста, подождите..."
    )
    
    # Получаем данные
    data = await state.get_data()
    tracks_count = data['tracks_count']
    all_tracks_list = data['all_tracks_list']
    total_tracks_count = data['total_tracks_count']
    orders_2d = data['orders_2d']
    optimization_result = data['optimization_result']
    plate_lookup_exact = data['plate_lookup_exact']
    plate_lookup_by_length = data['plate_lookup_by_length']
    
    # Восстанавливаем данные оптимизации
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
    
    # Вычисляем индексы
    start_index = (day_number - 1) * tracks_count
    end_index = min(day_number * tracks_count, total_tracks_count)
    tracks_for_this_day = end_index - start_index
    
    await callback.message.answer(
        f"📋 День {day_number}:\n"
        f"• Дорожки: {start_index + 1}-{end_index}\n"
        f"• Количество: {tracks_for_this_day} дорожек\n\n"
        f"⏳ Генерирую файлы..."
    )
    
    # Генерируем визуализацию
    result_paths = await asyncio.to_thread(
        visualize_plan,
        OUTPUTS_DIR_STR,
        tracks_for_this_day,
        start_index,
        use_production_pricing=True
    )
    
    if isinstance(result_paths, tuple) and len(result_paths) >= 2:
        png_path, pdf_schema_path = result_paths
        
        if os.path.exists(pdf_schema_path):
            await callback.message.answer_document(
                FSInputFile(pdf_schema_path),
                caption=f"📐 Схема дорожек {start_index + 1}-{end_index} (День {day_number})"
            )
        
        # Ищем детальную разбивку
        first_track = start_index + 1
        last_track = end_index
        
        if tracks_for_this_day == 1:
            breakdown_pattern_prefix = f'Детальная_разбивка_Дорожка_{first_track}_'
        else:
            breakdown_pattern_prefix = f'Детальная_разбивка_Дорожки_{first_track}-{last_track}_'
        
        breakdown_path = None
        try:
            for filename in os.listdir(OUTPUTS_DIR_STR):
                if filename.startswith(breakdown_pattern_prefix) and filename.endswith('.xlsx'):
                    candidate_path = os.path.join(OUTPUTS_DIR_STR, filename)
                    if breakdown_path is None or os.path.getctime(candidate_path) > os.path.getctime(breakdown_path):
                        breakdown_path = candidate_path
        except Exception as e:
            logger.exception(f"Ошибка поиска файла разбивки: {e}")
        
        if breakdown_path and os.path.exists(breakdown_path):
            await callback.message.answer_document(
                FSInputFile(breakdown_path),
                caption=f"📊 Детальная разбивка (День {day_number})"
            )
        
        # Формируем список плит
        tracks_in_current_file = all_tracks_list[start_index:end_index]
        
        # Список для хранения данных формовки по всем дорожкам
        formovka_tracks_data = []
        
        # Создаем КОПИЮ lookup для формовки (чтобы не влиять на оригинал в state)
        import copy
        formovka_lookup_exact = copy.deepcopy(plate_lookup_exact)
        formovka_lookup_by_length = copy.deepcopy(plate_lookup_by_length)
        
        def get_plate_info_smart(length, width):
            """
            Умный поиск информации о плите С УЧЕТОМ КОЛИЧЕСТВА.
            
            Логика:
            1. Ищем в списке записей по (length, width)
            2. Находим первую запись с qty_remaining > 0
            3. Уменьшаем qty_remaining на 1 (списываем плиту)
            4. Возвращаем информацию о КП
            
            ВАЖНО: Работаем с КОПИЕЙ lookup, чтобы не влиять на оригинал.
            """
            # 1. Сначала пробуем точное совпадение
            key = (round(length, 2), width)
            entries = formovka_lookup_exact.get(key, [])
            
            for entry in entries:
                if entry.get('qty_remaining', 0) > 0:
                    entry['qty_remaining'] -= 1
                    return entry.copy()
            
            # 2. Если ширина < 1200 (плита с резом), ищем по оригинальной ширине 1200
            if width < 1200:
                key_original = (round(length, 2), 1200)
                entries = formovka_lookup_exact.get(key_original, [])
                for entry in entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        return entry.copy()
            
            # 3. Fallback: поиск только по длине
            length_key = round(length, 2)
            entries = formovka_lookup_by_length.get(length_key, [])
            for entry in entries:
                if entry.get('qty_remaining', 0) > 0:
                    entry['qty_remaining'] -= 1
                    return entry.copy()
            
            return {
                'kp_date': 'неизвестно',
                'customer': 'неизвестно',
                'plate_name': '',
                'reinforcement': 0,
                'kp_id': None
            }
        
        for track_idx_in_file, track in enumerate(tracks_in_current_file):
            track_number = start_index + track_idx_in_file + 1
            track_items = track.get('items', [])
            
            if not track_items:
                continue
            
            plates_info = []
            for item in track_items:
                if item is None:
                    continue
                length = item.get('length')
                
                # Определяем ширину в зависимости от режима плиты
                mode = item.get('mode', 'solid')
                if mode == 'transverse' and item.get('width'):
                    width = int(item['width'] * 1000)  # width в метрах -> мм
                elif mode == 'split' and item.get('main_w'):
                    width = int(item['main_w'] * 1000)  # main_w в метрах -> мм
                else:
                    width = 1200  # solid или дефолт
                
                if not length:
                    continue
                
                plate_info = get_plate_info_smart(length, width)
                
                found = False
                for existing in plates_info:
                    if (round(existing['length'], 2) == round(length, 2) and
                        existing['width'] == width and
                        abs(existing['reinforcement'] - plate_info['reinforcement']) < 0.1 and
                        existing['kp_date'] == plate_info['kp_date'] and
                        existing['customer'] == plate_info['customer'] and
                        existing.get('plate_name', '') == plate_info.get('plate_name', '')):
                        existing['qty'] += 1
                        found = True
                        break
                
                if not found:
                    plates_info.append({
                        'length': length,
                        'width': width,
                        'qty': 1,
                        'reinforcement': plate_info['reinforcement'],
                        'kp_date': plate_info['kp_date'],
                        'customer': plate_info['customer'],
                        'plate_name': plate_info.get('plate_name', ''),
                        'kp_id': plate_info.get('kp_id')
                    })
                
                # НОВОЕ: Обрабатываем плиты из вторичных резов (остатков)
                secondary_cuts = item.get('secondary_cuts', []) if item else []
                for sec_cut in (secondary_cuts or []):
                    sec_width_m = sec_cut.get('width', 0)
                    if sec_width_m <= 0:
                        continue
                    
                    sec_width = int(sec_width_m * 1000)  # в мм
                    # Длина: если есть target_length (поперечный рез), иначе длина родительской плиты
                    sec_length = sec_cut.get('target_length') or length
                    
                    sec_plate_info = get_plate_info_smart(sec_length, sec_width)
                    
                    # Добавляем в plates_info (аналогично основной плите)
                    sec_found = False
                    for existing in plates_info:
                        if (round(existing['length'], 2) == round(sec_length, 2) and
                            existing['width'] == sec_width and
                            existing['kp_date'] == sec_plate_info.get('kp_date', 'неизвестно')):
                            existing['qty'] += 1
                            sec_found = True
                            break
                    
                    if not sec_found:
                        # Формируем имя плиты из label (если есть)
                        sec_plate_name = sec_plate_info.get('plate_name', '')
                        if not sec_plate_name and sec_cut.get('label'):
                            # Убираем префикс "О " из label
                            sec_plate_name = sec_cut['label'].replace('О ', '').strip()
                        
                        plates_info.append({
                            'length': sec_length,
                            'width': sec_width,
                            'qty': 1,
                            'reinforcement': sec_plate_info.get('reinforcement', 0),
                            'kp_date': sec_plate_info.get('kp_date', 'неизвестно'),
                            'customer': sec_plate_info.get('customer', 'неизвестно'),
                            'plate_name': sec_plate_name,
                            'kp_id': sec_plate_info.get('kp_id')
                        })
            
            if plates_info:
                # Получаем максимальное армирование дорожки
                max_reinforcement = track.get('max_reinforcement', 0)
                logger.debug(
                    f"[FORMOVKA] Дорожка {track_number}: {len(plates_info)} плит, макс. арм. {max_reinforcement}"
                )
                
                # Формируем заголовок с армированием дорожки
                if max_reinforcement > 0:
                    track_message = f"📋 Дорожка {track_number} (макс. арм. {max_reinforcement:.1f}):\n\n"
                else:
                    track_message = f"📋 Дорожка {track_number}:\n\n"
                
                plates_info.sort(key=lambda x: x['length'], reverse=True)
                
                for plate in plates_info:
                    plate_name = plate.get('plate_name', '')
                    
                    # Если есть готовое имя плиты из КП - используем его
                    if plate_name:
                        plate_str = f"{plate_name}"
                    else:
                        # Формируем имя плиты самостоятельно
                        length_dm = int(round(plate['length'] * 10))
                        width_mm = int(plate['width'])
                        
                        # Определяем нагрузку (по умолчанию 8п)
                        load_code = 8
                        if plate.get('reinforcement', 0) > 0:
                            # Примерное соответствие армирования и нагрузки
                            reinforcement = plate['reinforcement']
                            if reinforcement < 8:
                                load_code = 6
                            elif reinforcement < 12:
                                load_code = 8
                            elif reinforcement < 15:
                                load_code = 10
                            else:
                                load_code = 12
                        
                        # Форматируем ширину
                        if width_mm == 1200:
                            width_str = "12"
                        else:
                            width_dm = width_mm / 100.0
                            if abs(width_dm - int(width_dm)) < 0.01:
                                width_str = str(int(width_dm))
                            else:
                                width_str = str(width_dm).replace('.', ',')
                        
                        plate_str = f"ПБ {length_dm}-{width_str}-{load_code}п"
                        # Сохраняем сформированное имя обратно в plate_name для использования в Excel
                        plate['plate_name'] = plate_str
                    
                    track_message += (
                        f"  🔹 {plate_str} × {plate['qty']} шт "
                        f"(срок {plate['kp_date']}, заказчик: {plate['customer']})\n"
                    )
                
                # Сохраняем данные для файла формовки
                formovka_tracks_data.append({
                    'track_number': track_number,
                    'max_reinforcement': max_reinforcement,
                    'plates_info': plates_info
                })
                logger.debug(
                    f"[FORMOVKA] Добавлена дорожка {track_number} в formovka_tracks_data (всего: {len(formovka_tracks_data)})"
                )
                
                if len(track_message) > 4000:
                    lines = track_message.split('\n')
                    current_part = lines[0] + '\n\n'
                    
                    for line in lines[2:]:
                        if len(current_part + line + '\n') > 3900:
                            await callback.message.answer(current_part)
                            current_part = line + '\n'
                        else:
                            current_part += line + '\n'
                    
                    if current_part.strip():
                        await callback.message.answer(current_part)
                else:
                    await callback.message.answer(track_message)
        
        # Создаем Excel-файлы формовки для каждой дорожки
        logger.debug(f"[FORMOVKA] formovka_tracks_data count: {len(formovka_tracks_data)}")
        if formovka_tracks_data:
            logger.debug(f"[FORMOVKA] Начинаем создание {len(formovka_tracks_data)} файлов формовки")
            await callback.message.answer("📄 Генерирую файлы формовки...")
            
            # Создаем файлы формовки
            template_path = os.path.join(PROJECT_ROOT, "банк знаний", "!КЗ ПБ Шаблон.xlsx")
            logger.debug(f"[FORMOVKA] Template path: {template_path}")
            logger.debug(f"[FORMOVKA] Template exists: {os.path.exists(template_path)}")
            
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            try:
                formovka_files = await asyncio.to_thread(
                    create_formovka_files_for_tracks,
                    formovka_tracks_data,
                    OUTPUTS_DIR_STR,
                    template_path=template_path,
                    date_str=date_str
                )
                logger.debug(f"[FORMOVKA] Создано файлов: {len(formovka_files)}")
            except Exception as e:
                logger.exception(f"[FORMOVKA] Ошибка создания файлов: {e}")
                await callback.message.answer(
                    "⚠️ Не удалось создать файлы формовки.\n"
                    "Подробности в logs/bot.log."
                )
                formovka_files = []
            
            # Отправляем файлы формовки
            for formovka_file in formovka_files:
                if os.path.exists(formovka_file):
                    track_num = os.path.basename(formovka_file).split('_')[2]  # Извлекаем номер дорожки
                    await callback.message.answer_document(
                        FSInputFile(formovka_file),
                        caption=f"📋 Формовка (Дорожка {track_num})"
                    )
    
    await callback.message.answer(
        f"✅ План для Дня {day_number} готов!\n\n"
        f"Нажмите «День выполнен» когда закончите производство,\n"
        f"или выберите другой день:",
        reply_markup=production_day_actions_kb(day_number, data['total_days'])
    )
    
    await callback.answer()


@router.callback_query(F.data == "export_gantt")
async def export_gantt_chart(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "Диаграмма Ганта".
    Создаёт Excel-файл с визуализацией плана производства по КП.
    """
    await callback.message.answer("📊 Создаю диаграмму Ганта...")
    
    # Получаем данные из state
    data = await state.get_data()
    all_tracks_list = data.get('all_tracks_list', [])
    tracks_count = data.get('tracks_count', 1)
    plate_lookup_exact = data.get('plate_lookup_exact', {})
    plate_lookup_by_length = data.get('plate_lookup_by_length', {})
    total_days = data.get('total_days', 1)
    
    if not all_tracks_list:
        await callback.message.answer(
            "❌ Нет данных для создания диаграммы.\n"
            "Сначала выполните анализ производства.",
            reply_markup=main_menu_kb()
        )
        await callback.answer()
        return
    
    try:
        # Создаём диаграмму Ганта
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_list,
            tracks_count=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            output_dir=OUTPUTS_DIR_STR,
            start_date=datetime.now()
        )
        
        if gantt_path and os.path.exists(gantt_path):
            await callback.message.answer_document(
                FSInputFile(gantt_path),
                caption=(
                    "📊 Диаграмма Ганта\n\n"
                    "Цветовая кодировка:\n"
                    "🟢 Зелёный — успеваем до дедлайна\n"
                    "🟡 Жёлтый — завершаем в день дедлайна\n"
                    "🔴 Красный — опаздываем!"
                )
            )
        else:
            await callback.message.answer(
                "⚠️ Не удалось создать диаграмму.\n"
                "Возможно, нет данных о КП в плане производства."
            )
    
    except Exception as e:
        logger.exception(f"Ошибка создания диаграммы: {e}")
        await callback.message.answer(
            "❌ Не удалось создать диаграмму Ганта.\n"
            "Подробности в logs/bot.log."
        )
    
    # Показываем клавиатуру выбора дней снова
    await callback.message.answer(
        "Выберите день для просмотра:",
        reply_markup=production_days_kb(total_days)
    )
    
    await callback.answer()


@router.callback_query(F.data == "save_current_plan")
async def save_current_plan(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик сохранения актуального плана производства.
    
    Простыми словами:
    - Получает текущий план производства из state
    - Генерирует диаграмму Ганта
    - Сохраняет информацию о плане в JSON файл для дальнейшего использования
    - План можно будет просмотреть через Архив -> Актуальный план
    """
    await callback.message.answer("💾 Сохраняю актуальный план производства...")
    
    # Получаем данные из state
    data = await state.get_data()
    all_tracks_list = data.get('all_tracks_list', [])
    tracks_count = data.get('tracks_count', 1)
    plate_lookup_exact = data.get('plate_lookup_exact', {})
    plate_lookup_by_length = data.get('plate_lookup_by_length', {})
    total_days = data.get('total_days', 1)
    target_date = data.get('target_date', datetime.now().isoformat())
    
    if not all_tracks_list:
        await callback.message.answer(
            "❌ Нет данных для сохранения.\n"
            "Сначала выполните анализ производства.",
            reply_markup=main_menu_kb()
        )
        await callback.answer()
        return
    
    try:
        # Создаём диаграмму Ганта
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_list,
            tracks_count=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            output_dir=OUTPUTS_DIR_STR,
            start_date=datetime.now()
        )
        
        if not gantt_path or not os.path.exists(gantt_path):
            await callback.message.answer(
                "⚠️ Не удалось создать диаграмму Ганта.\n"
                "План не сохранён."
            )
            await callback.answer()
            return
        
        # Создаём директорию для данных, если её нет
        data_dir = BOT_DIR / "data"
        data_dir.mkdir(exist_ok=True)
        
        # Сохраняем информацию о плане в JSON
        current_plan_data = {
            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "gantt_file_path": gantt_path,
            "total_days": total_days,
            "tracks_count": tracks_count,
            "target_date": target_date,
            # Сохраняем данные для возможного восстановления плана
            "all_tracks_list": all_tracks_list,
            "plate_lookup_exact": {
                str(k): v for k, v in plate_lookup_exact.items()
            },
            "plate_lookup_by_length": {
                str(k): v for k, v in plate_lookup_by_length.items()
            }
        }
        
        json_path = data_dir / "current_plan.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(current_plan_data, f, ensure_ascii=False, indent=2)
        
        await callback.message.answer(
            f"✅ Актуальный план успешно сохранён!\n\n"
            f"📊 Информация о плане:\n"
            f"  • Дней производства: {total_days}\n"
            f"  • Дорожек в день: {tracks_count}\n"
            f"  • Всего дорожек: {len(all_tracks_list)}\n"
            f"  • Дата сохранения: {current_plan_data['saved_at']}\n\n"
            f"💡 Теперь вы можете просмотреть этот план через:\n"
            f"Архив → Актуальный план"
        )
        
    except Exception as e:
        logger.exception(f"Ошибка при сохранении плана: {e}")
        await callback.message.answer(
            "❌ Не удалось сохранить план.\n"
            "Подробности в logs/bot.log."
        )
    
    await callback.answer()


# === ОБРАБОТЧИКИ ЗАВЕРШЕНИЯ ДНЯ ===

@router.callback_query(F.data.startswith("complete_day_"))
async def start_day_completion(callback: CallbackQuery, state: FSMContext):
    """
    Начало процесса завершения дня.
    Показываем плиты для отметки брака.
    """
    day_number = int(callback.data.split("_")[-1])
    data = await state.get_data()
    
    tracks_count = data.get('tracks_count', 1)
    all_tracks_list = data.get('all_tracks_list', [])
    plate_lookup_exact = data.get('plate_lookup_exact', {})
    plate_lookup_by_length = data.get('plate_lookup_by_length', {})
    
    if not all_tracks_list:
        await callback.message.answer(
            "❌ Данные о дорожках не найдены. Попробуйте заново.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
        await callback.answer()
        return
    
    # Собираем все плиты дня
    start_index = (day_number - 1) * tracks_count
    end_index = min(day_number * tracks_count, len(all_tracks_list))
    tracks_for_day = all_tracks_list[start_index:end_index]
    
    # Создаем КОПИЮ lookup для завершения дня (чтобы не влиять на оригинал в state)
    import copy
    completion_lookup_exact = copy.deepcopy(plate_lookup_exact)
    completion_lookup_by_length = copy.deepcopy(plate_lookup_by_length)
    
    # Функция для получения информации о плите (с списанием из lookup)
    def get_plate_info_smart(length, width):
        """
        Умный поиск информации о плите С УЧЕТОМ КОЛИЧЕСТВА.
        
        Логика:
        1. Ищем в списке записей по (length, width)
        2. Находим первую запись с qty_remaining > 0
        3. Уменьшаем qty_remaining на 1 (списываем плиту)
        4. Возвращаем информацию о КП
        
        ВАЖНО: Работаем с КОПИЕЙ lookup, чтобы не влиять на оригинал.
        """
        # 1. Сначала пробуем точное совпадение
        key = (round(length, 2), width)
        entries = completion_lookup_exact.get(key, [])
        
        for entry in entries:
            if entry.get('qty_remaining', 0) > 0:
                entry['qty_remaining'] -= 1
                return entry.copy()
        
        # 2. Если ширина < 1200 (плита с резом), ищем по оригинальной ширине 1200
        if width < 1200:
            key_original = (round(length, 2), 1200)
            entries = completion_lookup_exact.get(key_original, [])
            for entry in entries:
                if entry.get('qty_remaining', 0) > 0:
                    entry['qty_remaining'] -= 1
                    return entry.copy()
        
        # 3. Fallback: поиск только по длине
        length_key = round(length, 2)
        entries = completion_lookup_by_length.get(length_key, [])
        for entry in entries:
            if entry.get('qty_remaining', 0) > 0:
                entry['qty_remaining'] -= 1
                return entry.copy()
        
        return {
            'kp_date': 'неизвестно',
            'customer': 'неизвестно',
            'plate_name': '',
            'kp_id': None
        }
    
    # Собираем плиты по дорожкам (каждая дорожка отдельно)
    day_plates_by_track = []
    total_qty = 0
    
    for track_idx, track in enumerate(tracks_for_day):
        track_number = start_index + track_idx + 1
        track_plates = []
        
        for item in track.get('items', []):
            if item is None:
                continue
            length = item.get('length', 0)
            
            # Определяем ширину в зависимости от режима плиты
            mode = item.get('mode', 'solid')
            if mode == 'transverse' and item.get('width'):
                width = int(item['width'] * 1000)  # width в метрах -> мм
            elif mode == 'split' and item.get('main_w'):
                width = int(item['main_w'] * 1000)  # main_w в метрах -> мм
            else:
                width = 1200  # solid или дефолт
            
            if not length:
                continue
            
            plate_info = get_plate_info_smart(length, width)
            plate_name = plate_info.get('plate_name', '')
            kp_id = plate_info.get('kp_id')
            
            # Если нет имени плиты — формируем его
            if not plate_name:
                length_dm = int(round(length * 10))
                width_mm = int(width)
                if width_mm == 1200:
                    width_str = "12"
                else:
                    width_dm = width_mm / 100.0
                    if abs(width_dm - int(width_dm)) < 0.01:
                        width_str = str(int(width_dm))
                    else:
                        width_str = str(width_dm).replace('.', ',')
                plate_name = f"ПБ {length_dm}-{width_str}-8п"
            
            # Получаем дату и заказчика для группировки
            kp_date = plate_info.get('kp_date', 'неизвестно')
            customer = plate_info.get('customer', 'неизвестно')
            
            # Ищем такую же плиту в списке текущей дорожки
            # Группируем по: plate_name + kp_id + kp_date + customer
            found = False
            for existing in track_plates:
                if (existing['plate_name'] == plate_name and 
                    existing['kp_id'] == kp_id and
                    existing['kp_date'] == kp_date and
                    existing['customer'] == customer):
                    existing['qty'] += 1
                    found = True
                    break
            
            if not found:
                track_plates.append({
                    'plate_name': plate_name,
                    'length_m': length,
                    'width_m': width / 1000.0,
                    'load_class': 800,
                    'qty': 1,
                    'kp_id': kp_id,
                    'kp_date': kp_date,
                    'customer': customer
                })
            
            # НОВОЕ: Обрабатываем плиты из вторичных резов (остатков)
            secondary_cuts = item.get('secondary_cuts', []) if item else []
            for sec_cut in (secondary_cuts or []):
                sec_width_m = sec_cut.get('width', 0)
                if sec_width_m <= 0:
                    continue
                
                sec_width = int(sec_width_m * 1000)  # в мм
                # Длина: если есть target_length (поперечный рез), иначе длина родительской плиты
                sec_length = sec_cut.get('target_length') or length
                
                sec_plate_info = get_plate_info_smart(sec_length, sec_width)
                sec_plate_name = sec_plate_info.get('plate_name', '')
                sec_kp_id = sec_plate_info.get('kp_id')
                
                # Если нет имени плиты — берём из label
                if not sec_plate_name and sec_cut.get('label'):
                    sec_plate_name = sec_cut['label'].replace('О ', '').strip()
                
                # Если всё ещё нет — формируем
                if not sec_plate_name:
                    sec_length_dm = int(round(sec_length * 10))
                    if sec_width == 1200:
                        sec_width_str = "12"
                    else:
                        sec_width_dm = sec_width / 100.0
                        if abs(sec_width_dm - int(sec_width_dm)) < 0.01:
                            sec_width_str = str(int(sec_width_dm))
                        else:
                            sec_width_str = str(sec_width_dm).replace('.', ',')
                    sec_plate_name = f"ПБ {sec_length_dm}-{sec_width_str}-8п"
                
                sec_kp_date = sec_plate_info.get('kp_date', 'неизвестно')
                sec_customer = sec_plate_info.get('customer', 'неизвестно')
                
                # Ищем такую же плиту в списке
                sec_found = False
                for existing in track_plates:
                    if (existing['plate_name'] == sec_plate_name and 
                        existing['kp_id'] == sec_kp_id and
                        existing['kp_date'] == sec_kp_date):
                        existing['qty'] += 1
                        sec_found = True
                        break
                
                if not sec_found:
                    track_plates.append({
                        'plate_name': sec_plate_name,
                        'length_m': sec_length,
                        'width_m': sec_width / 1000.0,
                        'load_class': 800,
                        'qty': 1,
                        'kp_id': sec_kp_id,
                        'kp_date': sec_kp_date,
                        'customer': sec_customer
                    })
        
        if track_plates:
            day_plates_by_track.append({
                'track_number': track_number,
                'plates': track_plates
            })
            total_qty += sum(p['qty'] for p in track_plates)
    
    # === ДОБАВЛЯЕМ ПЛИТЫ ИЗ ОСТАТКОВ (если есть) ===
    # Плиты из остатков добавляются как отдельная "дорожка" с номером 0
    plates_from_rests = data.get('plates_from_rests', [])
    rests_for_this_day = []  # Плиты из остатков для текущего дня
    
    if plates_from_rests and day_number == 1:
        # Плиты из остатков показываем только в День 1
        # (так как они не требуют производства на дорожках)
        for rest_plate in plates_from_rests:
            rests_for_this_day.append({
                'plate_name': rest_plate.get('plate_name', ''),
                'length_m': rest_plate.get('length_m', 0),
                'width_m': rest_plate.get('width_mm', 0) / 1000.0,
                'load_class': rest_plate.get('load_code', 8) * 100,
                'qty': rest_plate.get('qty', 1),
                'kp_id': rest_plate.get('kp_id'),
                'kp_date': rest_plate.get('kp_date', 'неизвестно'),
                'customer': rest_plate.get('customer', 'неизвестно'),
                # Флаги для отличия от обычных плит
                'from_rest': True,
                'rest_id': rest_plate.get('rest_id'),
                'match_type': rest_plate.get('match_type', 'exact'),
                'cut_cost': rest_plate.get('cut_cost', 0)
            })
        
        if rests_for_this_day:
            # Добавляем как "Дорожку 0: Из остатков" в начало списка
            day_plates_by_track.insert(0, {
                'track_number': 0,  # Номер 0 = из остатков
                'plates': rests_for_this_day,
                'is_from_rests': True  # Флаг для отображения в клавиатуре
            })
            total_qty += sum(p['qty'] for p in rests_for_this_day)
    
    if not day_plates_by_track:
        await callback.message.answer(
            f"❌ Не удалось найти плиты для Дня {day_number}.",
            reply_markup=production_days_kb(data.get('total_days', 1))
        )
        await callback.answer()
        return
    
    # Подсчитываем общее количество позиций
    total_positions = sum(len(track['plates']) for track in day_plates_by_track)
    
    # Сохраняем данные для завершения дня
    await state.update_data(
        completing_day=day_number,
        day_plates_by_track=day_plates_by_track,
        rejected_quantities={},  # Словарь {(track_idx, plate_idx): количество_брака}
        active_plate_id=None  # Какая плита сейчас открыта для редактирования
    )
    
    # Формируем сообщение
    await callback.message.answer(
        f"📋 День {day_number} — завершение производства\n\n"
        f"Всего плит: {total_qty} шт ({total_positions} позиций)\n"
        f"Дорожек: {len(day_plates_by_track)}\n\n"
        f"❗ Отметьте плиты, которые ушли в БРАК:\n"
        f"(Бракованные плиты останутся на следующий день)\n\n"
        f"Нажмите на плиту, чтобы выбрать количество брака:",
        reply_markup=plates_completion_kb(day_plates_by_track, {}, None)
    )
    
    await state.set_state(ProductionStates.marking_completion)
    await callback.answer()


@router.callback_query(F.data.startswith("plate_open_"), ProductionStates.marking_completion)
async def open_plate_counter(callback: CallbackQuery, state: FSMContext):
    """
    Открывает счетчик брака под плитой.
    Позволяет выбрать количество бракованных плит кнопками +/-.
    """
    # Парсим формат: plate_open_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    
    # Устанавливаем активную плиту
    active_plate_id = (track_idx, plate_idx)
    await state.update_data(active_plate_id=active_plate_id)
    
    # Обновляем клавиатуру
    await callback.message.edit_reply_markup(
        reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, active_plate_id)
    )
    await callback.answer()


@router.callback_query(F.data.startswith("reject_plus_"), ProductionStates.marking_completion)
async def increase_rejection(callback: CallbackQuery, state: FSMContext):
    """
    Увеличивает количество брака на 1.
    Максимум = общему количеству плит.
    """
    # Парсим формат: reject_plus_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    active_plate_id = data.get('active_plate_id')
    
    # Получаем максимальное количество плит
    plate = day_plates_by_track[track_idx]['plates'][plate_idx]
    max_qty = plate.get('qty', 1)
    
    # Увеличиваем брак (макс = max_qty)
    plate_id = (track_idx, plate_idx)
    current = rejected_quantities.get(plate_id, 0)
    if current < max_qty:
        rejected_quantities[plate_id] = current + 1
        await state.update_data(rejected_quantities=rejected_quantities)
        
        # Обновляем клавиатуру
        await callback.message.edit_reply_markup(
            reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, active_plate_id)
        )
    await callback.answer()


@router.callback_query(F.data.startswith("reject_minus_"), ProductionStates.marking_completion)
async def decrease_rejection(callback: CallbackQuery, state: FSMContext):
    """
    Уменьшает количество брака на 1.
    Минимум = 0.
    """
    # Парсим формат: reject_minus_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    active_plate_id = data.get('active_plate_id')
    
    # Уменьшаем брак (мин = 0)
    plate_id = (track_idx, plate_idx)
    current = rejected_quantities.get(plate_id, 0)
    if current > 0:
        rejected_quantities[plate_id] = current - 1
        # Если стало 0, удаляем из словаря
        if rejected_quantities[plate_id] == 0:
            del rejected_quantities[plate_id]
        await state.update_data(rejected_quantities=rejected_quantities)
        
        # Обновляем клавиатуру
        await callback.message.edit_reply_markup(
            reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, active_plate_id)
        )
    await callback.answer()


@router.callback_query(F.data.startswith("reject_reset_"), ProductionStates.marking_completion)
async def reset_rejection(callback: CallbackQuery, state: FSMContext):
    """
    Сбрасывает брак в 0 и закрывает счетчик.
    """
    # Парсим формат: reject_reset_t{track_idx}_p{plate_idx}
    parts = callback.data.split("_")
    track_idx = int(parts[2][1:])  # Убираем 't' из 't0'
    plate_idx = int(parts[3][1:])  # Убираем 'p' из 'p0'
    
    data = await state.get_data()
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    
    # Сбрасываем брак и закрываем счетчик
    plate_id = (track_idx, plate_idx)
    if plate_id in rejected_quantities:
        del rejected_quantities[plate_id]
    
    await state.update_data(
        rejected_quantities=rejected_quantities,
        active_plate_id=None
    )
    
    # Обновляем клавиатуру
    await callback.message.edit_reply_markup(
        reply_markup=plates_completion_kb(day_plates_by_track, rejected_quantities, None)
    )
    await callback.answer("Брак сброшен")


@router.callback_query(F.data.startswith("reject_info_"), ProductionStates.marking_completion)
async def reject_info_click(callback: CallbackQuery):
    """
    Обработчик для центральной кнопки 'Брак: X/Y' (без действия).
    """
    await callback.answer()


@router.callback_query(F.data.startswith("track_header_"), ProductionStates.marking_completion)
async def track_header_click(callback: CallbackQuery):
    """
    Обработчик для заголовков дорожек.
    Не делает ничего, просто отвечает на callback.
    """
    await callback.answer("Это заголовок дорожки")


@router.callback_query(F.data == "confirm_completion", ProductionStates.marking_completion)
async def confirm_day_completion(callback: CallbackQuery, state: FSMContext):
    """
    Подтверждение завершения дня.
    Переносит выполненные плиты в completed_plates.
    Сохраняет информацию об остатках в plate_rests.
    """
    data = await state.get_data()
    day_number = data.get('completing_day', 1)
    day_plates_by_track = data.get('day_plates_by_track', [])
    rejected_quantities = data.get('rejected_quantities', {})
    
    # Разделяем на выполненные и бракованные с учетом частичного брака
    completed_plates = []
    rejected_plates = []
    
    for track_idx, track_data in enumerate(day_plates_by_track):
        plates = track_data.get('plates', [])
        track_number = track_data.get('track_number', track_idx + 1)
        
        for plate_idx, plate in enumerate(plates):
            plate_id = (track_idx, plate_idx)
            reject_qty = rejected_quantities.get(plate_id, 0)
            total_qty = plate.get('qty', 1)
            
            if reject_qty == 0:
                # Все плиты выполнены
                completed_plates.append(plate)
            elif reject_qty >= total_qty:
                # Все плиты в браке
                plate_with_track = plate.copy()
                plate_with_track['track_number'] = track_number
                rejected_plates.append(plate_with_track)
            else:
                # ЧАСТИЧНЫЙ БРАК: разделяем плиту на 2 части
                # 1) Выполненная часть
                completed_part = plate.copy()
                completed_part['qty'] = total_qty - reject_qty
                completed_plates.append(completed_part)
                
                # 2) Бракованная часть
                rejected_part = plate.copy()
                rejected_part['qty'] = reject_qty
                rejected_part['track_number'] = track_number
                rejected_plates.append(rejected_part)
    
    # Импортируем функции БД
    db_path = str(PROJECT_ROOT / "plita.db")
    
    # Группируем по kp_id
    plates_by_kp = defaultdict(list)
    plates_without_kp = []  # Плиты без kp_id (не найдены в lookup)
    
    for plate in completed_plates:
        kp_id = plate.get('kp_id')
        if kp_id:
            plates_by_kp[kp_id].append(plate)
        else:
            plates_without_kp.append(plate)
    
    total_moved = 0
    completed_kps = []
    
    # === ОБРАБОТКА ПЛИТ ИЗ ОСТАТКОВ ===
    # Сначала обрабатываем плиты из остатков (помечаем остатки как использованные)
    rests_used_count = 0
    for plate in completed_plates:
        if plate.get('from_rest') and plate.get('rest_id'):
            rest_id = plate['rest_id']
            # Помечаем остаток как использованный
            if kp_db.mark_rest_as_used(rest_id, db_path):
                rests_used_count += 1
                logger.info(f"[COMPLETION] Остаток #{rest_id} помечен как использованный")
            
            # Переносим плиту в completed_plates
            kp_id = plate.get('kp_id')
            if kp_id:
                moved = kp_db.move_plates_to_completed(kp_id, [plate], day_number, db_path)
                total_moved += moved
                
                if kp_db.check_and_update_kp_completion(kp_id, db_path):
                    if kp_id not in completed_kps:
                        completed_kps.append(kp_id)
    
    # Переносим плиты С kp_id (стандартная логика, исключая плиты из остатков)
    for kp_id, plates in plates_by_kp.items():
        # Фильтруем плиты из остатков (они уже обработаны выше)
        plates_not_from_rests = [p for p in plates if not p.get('from_rest')]
        if not plates_not_from_rests:
            continue
        
        moved = kp_db.move_plates_to_completed(kp_id, plates_not_from_rests, day_number, db_path)
        total_moved += moved
        
        # Проверяем, завершён ли КП полностью
        if kp_db.check_and_update_kp_completion(kp_id, db_path):
            if kp_id not in completed_kps:
                completed_kps.append(kp_id)
    
    # ========== НОВОЕ: Обрабатываем плиты БЕЗ kp_id ==========
    # Эти плиты не были найдены в lookup-таблицах (возможно из-за изменения ширины после реза)
    # Ищем их в БД по длине
    import sqlite3
    for plate in plates_without_kp:
        length_m = plate.get('length_m', 0)
        plate_name = plate.get('plate_name', '')
        
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        
        # Ищем по длине среди КП "в работе"
        cur.execute('''
            SELECT p.kp_id, p.plate_name 
            FROM kp_plates p
            JOIN kp_meta m ON p.kp_id = m.kp_id
            WHERE ABS(p.length_m - ?) < 0.02 
              AND p.qty > 0
              AND m.status = 'в работе'
            ORDER BY p.kp_id
            LIMIT 1
        ''', (length_m,))
        row = cur.fetchone()
        conn.close()
        
        if row:
            found_kp_id = row[0]
            db_plate_name = row[1]
            plate['plate_name'] = db_plate_name  # Используем имя из БД для корректного списания
            
            moved = kp_db.move_plates_to_completed(found_kp_id, [plate], day_number, db_path)
            total_moved += moved
            logger.info(f"[COMPLETION] Плита найдена по длине: {plate_name} ({length_m}м) → КП #{found_kp_id}")
            
            if kp_db.check_and_update_kp_completion(found_kp_id, db_path):
                if found_kp_id not in completed_kps:
                    completed_kps.append(found_kp_id)
        else:
            logger.warning(f"[COMPLETION] Плита не найдена в БД: {plate_name} ({length_m}м)")
    # ========== КОНЕЦ НОВОЙ ЛОГИКИ ==========
    
    # ========== СОХРАНЕНИЕ ОСТАТКОВ ==========
    # Получаем метаданные об остатках из результата оптимизации
    optimization_result = data.get('optimization_result', {})
    rests_created = optimization_result.get('rests_created', [])
    rests_used = optimization_result.get('rests_used', [])
    
    # Подсчитываем неиспользованные остатки
    unused_rests_count = 0
    
    # Для каждого созданного остатка проверяем, использован ли он
    for rest in rests_created:
        is_used = any(
            abs(r['source_length'] - rest['length']) < 0.01 and 
            r['source_rest_mm'] == rest['rest_width_mm']
            for r in rests_used
        )
        
        if not is_used:
            # Остаток не использован - сохраняем в plate_rests
            # Находим kp_id для этого остатка через plate_lookup
            plate_lookup_exact = data.get('plate_lookup_exact', {})
            plate_lookup_by_length = data.get('plate_lookup_by_length', {})
            
            # Ищем kp_id по длине и ширине
            key = (round(rest['length'], 2), rest['source_width_mm'])
            plate_info_list = plate_lookup_exact.get(key, [])
            # plate_lookup_exact тоже возвращает СПИСОК, берём первый элемент
            plate_info = plate_info_list[0] if plate_info_list else None
            
            if not plate_info:
                length_key = round(rest['length'], 2)
                plate_info_list = plate_lookup_by_length.get(length_key, [])
                # plate_lookup_by_length возвращает СПИСОК, берём первый элемент
                plate_info = plate_info_list[0] if plate_info_list else None
            
            if plate_info and plate_info.get('kp_id'):
                kp_id = plate_info['kp_id']
                source_plate_name = plate_info.get('plate_name', f"ПБ {rest['length']}м-{rest['source_width_mm']}мм")
                
                kp_db.create_plate_rest(
                    kp_id=kp_id,
                    source_plate_name=source_plate_name,
                    rest_width_mm=rest['rest_width_mm'],
                    length_m=rest['length'],
                    production_day=day_number,
                    db_path=db_path
                )
                unused_rests_count += 1
    # ========== КОНЕЦ СОХРАНЕНИЯ ОСТАТКОВ ==========
    
    # Формируем отчёт
    report = f"✅ День {day_number} завершён!\n\n"
    report += f"📦 Выполнено плит: {total_moved} шт\n"
    
    # Информация о плитах из остатков
    if rests_used_count > 0:
        report += f"💰 Из остатков: {rests_used_count} шт (чистая прибыль!)\n"
    
    if rejected_plates:
        rejected_qty = sum(p['qty'] for p in rejected_plates)
        report += f"\n❌ В браке: {rejected_qty} шт ({len(rejected_plates)} позиций)\n"
        report += "Эти плиты останутся на следующий день:\n"
        
        # Группируем бракованные плиты по дорожкам для отчета
        rejected_by_track = defaultdict(list)
        for plate in rejected_plates:
            track_num = plate.get('track_number', '?')
            rejected_by_track[track_num].append(plate)
        
        for track_num in sorted(rejected_by_track.keys()):
            report += f"\n  Дорожка {track_num}:\n"
            for plate in rejected_by_track[track_num]:
                report += f"   • {plate['plate_name']} × {plate['qty']}\n"
    
    # Информация об остатках
    if unused_rests_count > 0:
        report += f"\n📦 Остатков на складе: {unused_rests_count} шт\n"
        report += "(Доступны для следующих заказов)"
    
    if completed_kps:
        report += f"\n🎉 Полностью выполнены КП: {', '.join(map(str, completed_kps))}\n"
        report += "Статус этих КП изменён на «выполнено»"
    
    await callback.message.answer(report, reply_markup=main_menu_kb())
    await state.clear()
    await callback.answer("✅ День завершён!")


@router.callback_query(F.data == "cancel_completion", ProductionStates.marking_completion)
async def cancel_day_completion(callback: CallbackQuery, state: FSMContext):
    """
    Отмена завершения дня.
    Возвращаемся к выбору дней.
    """
    data = await state.get_data()
    total_days = data.get('total_days', 1)
    
    # Очищаем данные завершения, но оставляем остальные данные
    await state.update_data(
        completing_day=None,
        day_plates_by_track=None,
        rejected_quantities=None,
        active_plate_id=None
    )
    
    await callback.message.answer(
        "❌ Завершение дня отменено.\n\n"
        "Выберите день для просмотра:",
        reply_markup=production_days_kb(total_days)
    )
    await state.set_state(ProductionStates.waiting_day_selection)
    await callback.answer()


@router.callback_query(F.data == "cancel_process")
async def cancel_production_process(callback: CallbackQuery, state: FSMContext):
    """Отмена процесса планирования производства"""
    await state.clear()
    await callback.message.answer(
        "❌ Планирование производства отменено.\n"
        "Выберите действие:",
        reply_markup=main_menu_kb()
    )
    await callback.answer("Отменено")

