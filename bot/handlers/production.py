"""Обработчики планирования производства плит"""
import asyncio
import logging
import os
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
import core.config_and_data as cfg
import core.optimization as optimization

from ..keyboards import main_menu_kb, production_days_kb
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
            return
    except ValueError:
        await message.answer(
            "❌ Неверный формат. Введите целое число (например: 5):"
        )
        return
    
    await state.update_data(tracks_count=tracks_count)
    
    await state.set_state(ProductionStates.waiting_date_number)
    await message.answer(
        f"✅ Дорожек: {tracks_count}\n\n"
        "Шаг 2 из 2: До какой даты брать плиты?\n\n"
        "Поддерживаемые форматы:\n"
        "• 25 (число текущего месяца)\n"
        "• 01.02.2026 (полная дата)\n"
        "• 2026-02-01 (ISO формат)\n\n"
        "Введите дату:"
    )


@router.message(ProductionStates.waiting_date_number)
async def receive_date_number_and_plan(message: Message, state: FSMContext):
    """Получаем дату, запускаем анализ и показываем кнопки выбора дня"""
    user_input = message.text.strip()
    
    # === ПАРСИНГ ДАТЫ: поддерживаем разные форматы ===
    target_date = None
    
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
        return
    
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    
    await message.answer(
        f"✅ Параметры планирования:\n"
        f"• Дорожек в день: {tracks_count}\n"
        f"• Плиты со сроком до: {date_description}\n\n"
        f"⏳ Загружаю плиты из базы данных..."
    )
    
    try:
        # === ШАГ 1-3: ЗАГРУЗКА И ПОДГОТОВКА ДАННЫХ ===
        db_path = PROJECT_ROOT / "plita.db"
        pb_db_path = BOT_DIR / "pb.db"
        
        import sqlite3
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        
        cur.execute("""
            SELECT kp.kp_id, kp.execution_terms, kp.customer_name
            FROM KP_offers kp
            JOIN kp_meta meta ON kp.kp_id = meta.kp_id
            WHERE meta.status = 'в работе'
        """)
        
        kp_list = []
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
        
        if not kp_list:
            conn.close()
            await message.answer(
                f"❌ Нет КП в работе со сроком до {date_description}.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        kp_list.sort(key=lambda x: x['date'])
        
        await message.answer(f"✅ Найдено КП в работе: {len(kp_list)}\nЗагружаю плиты...")
        
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
                'plate_name': plate_data.get('plate_name', '')
            })
        
        # Lookup-таблицы
        plate_lookup_exact = {}
        plate_lookup_by_length = {}
        
        for order in orders_2d:
            key = (round(order['length'], 2), order['width'])
            if key not in plate_lookup_exact:
                plate_lookup_exact[key] = {
                    'kp_date': order.get('kp_date', 'неизвестно'),
                    'customer': order.get('customer', 'неизвестно'),
                    'plate_name': order.get('plate_name', ''),
                    'reinforcement': order.get('reinforcement', 0),
                    'load_code': order.get('load_code', 8),
                    'qty': order.get('qty', 1),
                }
            
            length_key = round(order['length'], 2)
            if length_key not in plate_lookup_by_length:
                plate_lookup_by_length[length_key] = {
                    'kp_date': order.get('kp_date', 'неизвестно'),
                    'customer': order.get('customer', 'неизвестно'),
                    'plate_name': order.get('plate_name', ''),
                    'reinforcement': order.get('reinforcement', 0),
                }
        
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
                
                for i, item in enumerate(items):
                    item_length = item['length']
                    will_exceed = (current_track_length + item_length > MAX_TRACK_LENGTH and current_track)
                    
                    if will_exceed:
                        all_tracks_list.append({
                            'items': current_track,
                            'length': current_track_length,
                            'load_code': load_code,
                            'label': group_label
                        })
                        current_track = []
                        current_track_length = 0.0
                    
                    current_track.append(item)
                    current_track_length += item_length
                
                if current_track:
                    all_tracks_list.append({
                        'items': current_track,
                        'length': current_track_length,
                        'load_code': load_code,
                        'label': group_label
                    })
        else:
            current_track = []
            current_track_length = 0.0
            
            for i, item in enumerate(seq):
                item_length = item['length']
                will_exceed = (current_track_length + item_length > MAX_TRACK_LENGTH and current_track)
                
                if will_exceed:
                    all_tracks_list.append({
                        'items': current_track,
                        'length': current_track_length
                    })
                    current_track = []
                    current_track_length = 0.0
                
                current_track.append(item)
                current_track_length += item_length
            
            if current_track:
                all_tracks_list.append({
                    'items': current_track,
                    'length': current_track_length
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
            target_date=target_date.isoformat()
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
            "❌ Ошибка при планировании производства.\n"
            "Подробности в logs/bot.log.",
            reply_markup=main_menu_kb()
        )
        await state.clear()


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
        
        def get_plate_info_smart(length, width):
            key = (round(length, 2), width)
            info = plate_lookup_exact.get(key)
            if info:
                return info.copy()
            
            length_key = round(length, 2)
            info = plate_lookup_by_length.get(length_key)
            if info:
                return info.copy()
            
            return {
                'kp_date': 'неизвестно',
                'customer': 'неизвестно',
                'plate_name': '',
                'reinforcement': 0
            }
        
        for track_idx_in_file, track in enumerate(tracks_in_current_file):
            track_number = start_index + track_idx_in_file + 1
            track_items = track.get('items', [])
            
            if not track_items:
                continue
            
            plates_info = []
            for item in track_items:
                length = item.get('length')
                width = item.get('width', 1200)
                
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
                        'plate_name': plate_info.get('plate_name', '')
                    })
            
            if plates_info:
                track_message = f"📋 Дорожка {track_number}:\n\n"
                plates_info.sort(key=lambda x: x['length'], reverse=True)
                
                for plate in plates_info:
                    reinforcement_str = f"{plate['reinforcement']:.1f}"
                    plate_name = plate.get('plate_name', '')
                    
                    if plate_name:
                        plate_info_str = f"Плиты {plate_name}"
                    else:
                        length_str = f"{plate['length']:.2f}".replace('.', ',')
                        width_str = f"{plate['width']}" if plate['width'] != 1200 else "1200"
                        plate_info_str = f"Плиты {length_str}м × {width_str}мм"
                    
                    track_message += (
                        f"  📦 {plate_info_str} × {plate['qty']} шт "
                        f"(армир. {reinforcement_str}, срок {plate['kp_date']}, "
                        f"заказчик: {plate['customer']})\n"
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
    
    await callback.message.answer(
        f"✅ План для Дня {day_number} готов!\n\n"
        f"Хотите просмотреть другой день?",
        reply_markup=production_days_kb(data['total_days'])
    )
    
    await callback.answer()


@router.callback_query(F.data == "production_all_days")
async def process_all_days_selection(callback: CallbackQuery, state: FSMContext):
    """Обработчик выбора 'Все дни сразу'"""
    
    await callback.message.answer(
        f"📦 Генерирую планы для всех дней...\n"
        f"⏳ Это может занять некоторое время..."
    )
    
    data = await state.get_data()
    total_days = data['total_days']
    
    # Генерируем для каждого дня
    for day in range(1, total_days + 1):
        class FakeCallbackQuery:
            def __init__(self, msg, day_num):
                self.message = msg
                self.data = f"production_day_{day_num}"
            
            async def answer(self):
                pass
        
        fake_callback = FakeCallbackQuery(callback.message, day)
        await process_day_selection(fake_callback, state)
    
    await callback.message.answer(
        "✅ Все планы готовы!",
        reply_markup=main_menu_kb()
    )
    
    await state.clear()
    await callback.answer()

