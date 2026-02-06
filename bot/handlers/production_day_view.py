"""Просмотр дня производства и генерация документов"""
import asyncio
import logging
import os
import copy
from datetime import datetime, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import CallbackQuery, FSInputFile
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.visualization import visualize_plan
from core.formovka_excel import create_formovka_files_for_tracks
import core.config_and_data as cfg
import core.optimization as optimization

from ..keyboards import production_day_actions_kb, day_documents_menu_kb
from ..bot_config import OUTPUTS_DIR_STR

# Импорт функции для работы с мультипланами
from .plan_manager import get_tracks_for_date_from_all_plans

router = Router()


async def _restore_optimization_data(state: FSMContext, day_number: int):
    """
    Восстанавливает данные оптимизации из state для генерации документов.
    
    Простыми словами:
    - Загружает данные из памяти бота (state)
    - Восстанавливает настройки оптимизации
    - Вычисляет индексы дорожек для выбранного дня
    - Возвращает всё необходимое для генерации документов
    
    Args:
        state: состояние FSM (память бота)
        day_number: номер дня (например, 3)
    
    Returns:
        dict: словарь с данными для генерации документов
    """
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
    
    # Вычисляем индексы дорожек
    start_index = (day_number - 1) * tracks_count
    end_index = min(day_number * tracks_count, total_tracks_count)
    tracks_for_this_day = end_index - start_index
    
    return {
        'tracks_count': tracks_count,
        'all_tracks_list': all_tracks_list,
        'start_index': start_index,
        'end_index': end_index,
        'tracks_for_this_day': tracks_for_this_day,
        'orders_2d': orders_2d,
        'plate_lookup_exact': plate_lookup_exact,
        'plate_lookup_by_length': plate_lookup_by_length,
        'total_tracks_count': total_tracks_count
    }


def _restore_optimization_globals(orders_2d: list, optimization_result: dict):
    """
    Восстанавливает глобальные переменные оптимизации для генерации документов.
    
    Простыми словами:
    - Устанавливает глобальные переменные optimization.* и cfg.*
    - Эти переменные нужны для работы visualize_plan() и функций расчета себестоимости
    - При работе с мультипланами эти переменные не восстанавливаются автоматически
    
    Args:
        orders_2d: Список заказов (плит) для производства
        optimization_result: Результат оптимизации раскладки
    """
    # Восстанавливаем данные оптимизации
    optimization.OPT_CASCADING_PLAN = optimization_result
    
    all_loads = set(p['load_code'] for p in orders_2d) if orders_2d else {8}
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
    
    logger.debug(f"[RESTORE_GLOBALS] Восстановлено глобальных переменных: "
                f"loads={len(all_loads)}, PLATE_LOAD_DETAILS={len(cfg.PLATE_LOAD_DETAILS)}, "
                f"PLATES_1_2={len(cfg.PLATES_1_2)}")


@router.callback_query(F.data.startswith("production_day_"))
async def process_day_selection(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик выбора конкретного дня.
    
    Простыми словами:
    - Показывает текстовый состав дорожек (какие плиты производятся)
    - НЕ генерирует файлы сразу
    - Показывает меню выбора типа документа
    
    НОВАЯ ЛОГИКА:
    - Если from_saved_plan=True: собирает дорожки из ВСЕХ планов на эту дату
    - Если from_saved_plan=False: использует данные из state (старая логика)
    """
    
    day_number = int(callback.data.split("_")[-1])
    
    data = await state.get_data()
    from_saved_plan = data.get('from_saved_plan', False)
    
    # === НОВАЯ ЛОГИКА: Для сохранённых планов собираем из всех планов ===
    if from_saved_plan:
        # Вычисляем дату выбранного дня
        plan_start_date = data.get('plan_start_date')
        if not plan_start_date:
            await callback.message.answer("❌ Ошибка: не найдена дата начала плана")
            await callback.answer()
            return
        
        try:
            start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
        except ValueError:
            await callback.message.answer("❌ Ошибка: неверный формат даты начала плана")
            await callback.answer()
            return
        
        selected_date = (start_dt + timedelta(days=day_number - 1)).strftime('%Y-%m-%d')
        
        # Загружаем дорожки из всех планов на эту дату
        multi_plan_data = get_tracks_for_date_from_all_plans(selected_date)
        
        if not multi_plan_data:
            await callback.message.answer(
                f"❌ Дата {datetime.strptime(selected_date, '%Y-%m-%d').strftime('%d.%m.%Y')} "
                f"не найдена ни в одном сохранённом плане."
            )
            await callback.answer()
            return
        
        # Используем данные из всех планов
        tracks_in_current_file = multi_plan_data['tracks']
        tracks_for_this_day = len(tracks_in_current_file)
        plate_lookup_exact = multi_plan_data['plate_lookup_exact']
        plate_lookup_by_length = multi_plan_data['plate_lookup_by_length']
        
        # Нумерация дорожек: начинаем с 1
        start_index = 0
        end_index = tracks_for_this_day
        
        # Информационное сообщение о том, из скольких планов собраны дорожки
        plans_info = f" (из {multi_plan_data['plans_count']} планов)" if multi_plan_data['plans_count'] > 1 else ""
        
        logger.info(f"[MULTI_PLAN_VIEW] День {day_number} ({selected_date}): "
                   f"{tracks_for_this_day} дорожек из {multi_plan_data['plans_count']} планов")
        
        # НОВОЕ: Формируем детальную информацию о планах
        plans_detail = ""
        if multi_plan_data['plans_count'] > 1:
            from collections import Counter
            tracks_by_plan = Counter()
            
            # Подсчитываем дорожки по планам
            for track in tracks_in_current_file:
                if isinstance(track, dict) and 'source_plan_name' in track:
                    tracks_by_plan[track['source_plan_name']] += 1
            
            if tracks_by_plan:
                plans_detail = "\n\n📋 Дорожки из планов:\n"
                for plan_name, count in sorted(tracks_by_plan.items()):
                    plans_detail += f"  • {plan_name}: {count} дор.\n"
        
        await callback.message.answer(
            f"📋 День {day_number} ({datetime.strptime(selected_date, '%Y-%m-%d').strftime('%d.%m')}):\n"
            f"• Дорожки: {start_index + 1}-{end_index}\n"
            f"• Количество: {tracks_for_this_day} дорожек{plans_info}"
            f"{plans_detail}"
        )
    else:
        # === СТАРАЯ ЛОГИКА: Для несохранённых планов используем state ===
        restored_data = await _restore_optimization_data(state, day_number)
        
        start_index = restored_data['start_index']
        end_index = restored_data['end_index']
        tracks_for_this_day = restored_data['tracks_for_this_day']
        all_tracks_list = restored_data['all_tracks_list']
        plate_lookup_exact = restored_data['plate_lookup_exact']
        plate_lookup_by_length = restored_data['plate_lookup_by_length']
        
        tracks_in_current_file = all_tracks_list[start_index:end_index]
        
        await callback.message.answer(
            f"📋 День {day_number}:\n"
            f"• Дорожки: {start_index + 1}-{end_index}\n"
            f"• Количество: {tracks_for_this_day} дорожек"
        )
    
    # Создаем КОПИЮ lookup для формовки (чтобы не влиять на оригинал в state)
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
        
        FUZZY-ПОИСК: Если точный ключ не найден, ищем с tolerance 0.03м (30мм)
        по длине. Это нужно, потому что оптимизатор может округлять длины
        (например, 3.8м -> 3.79м или 5.71м -> 5.7м).
        """
        TOLERANCE = 0.03  # 30мм tolerance для fuzzy-поиска
        rounded_length = round(length, 2)
        
        # 1. Сначала пробуем точное совпадение
        key = (rounded_length, width)
        entries = formovka_lookup_exact.get(key, [])
        
        for entry in entries:
            if entry.get('qty_remaining', 0) > 0:
                entry['qty_remaining'] -= 1
                return entry.copy()
        
        # 2. Если ширина < 1200 (плита с резом), ищем по оригинальной ширине 1200
        if width < 1200:
            key_original = (rounded_length, 1200)
            entries = formovka_lookup_exact.get(key_original, [])
            for entry in entries:
                if entry.get('qty_remaining', 0) > 0:
                    entry['qty_remaining'] -= 1
                    return entry.copy()
        
        # 3. Fuzzy-поиск с tolerance по длине в exact lookup
        for lookup_key, entries in formovka_lookup_exact.items():
            key_length, key_width = lookup_key
            # Проверяем ширину (точно или 1200 для split)
            if key_width != width and key_width != 1200:
                continue
            # Проверяем длину с tolerance
            if abs(key_length - rounded_length) <= TOLERANCE:
                for entry in entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        return entry.copy()
        
        # 4. Fallback: поиск только по длине (точный)
        entries = formovka_lookup_by_length.get(rounded_length, [])
        for entry in entries:
            if entry.get('qty_remaining', 0) > 0:
                entry['qty_remaining'] -= 1
                return entry.copy()
        
        # 5. Fuzzy fallback: поиск по длине с tolerance
        for lookup_length, entries in formovka_lookup_by_length.items():
            if abs(lookup_length - rounded_length) <= TOLERANCE:
                for entry in entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        return entry.copy()
        
        return {
            'kp_id': None,
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
            if item is None:
                continue
            length = item.get('length')
            
            # Определяем ширину в зависимости от режима плиты
            mode = item.get('mode', 'solid')
            if mode == 'transverse' and item.get('width'):
                width = round(item['width'] * 1000)  # round для корректного округления float
            elif mode == 'split' and item.get('main_w'):
                width = round(item['main_w'] * 1000)  # round для корректного округления
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
                    existing.get('kp_id') == plate_info.get('kp_id') and
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
                    'kp_id': plate_info.get('kp_id'),
                    'plate_name': plate_info.get('plate_name', '')
                })
            
            # НОВОЕ: Обрабатываем плиты из вторичных резов (остатков)
            secondary_cuts = item.get('secondary_cuts', []) if item else []
            for sec_cut in (secondary_cuts or []):
                sec_width_m = sec_cut.get('width', 0)
                if sec_width_m <= 0:
                    continue
                
                sec_width = round(sec_width_m * 1000)  # round для корректного округления float
                # Длина: если есть target_length (поперечный рез), иначе длина родительской плиты
                sec_length = sec_cut.get('target_length') or length
                
                sec_plate_info = get_plate_info_smart(sec_length, sec_width)
                
                # Добавляем в plates_info (аналогично основной плите)
                sec_found = False
                for existing in plates_info:
                    if (round(existing['length'], 2) == round(sec_length, 2) and
                        existing['width'] == sec_width and
                        abs(existing['reinforcement'] - sec_plate_info.get('reinforcement', 0)) < 0.1 and
                        existing.get('kp_id') == sec_plate_info.get('kp_id') and
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
                        'kp_id': sec_plate_info.get('kp_id'),
                        'plate_name': sec_plate_name
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
    
    # Показываем меню выбора документов
    track_numbers_str = f"{start_index + 1}-{end_index}" if tracks_for_this_day > 1 else str(start_index + 1)
    
    await callback.message.answer(
        "📄 Выберите тип документа для генерации:",
        reply_markup=day_documents_menu_kb(day_number, track_numbers_str)
    )
    
    # Сохраняем данные в state для обработчиков генерации
    # ВАЖНО: При from_saved_plan=True сохраняем данные из мультипланов
    if from_saved_plan:
        await state.update_data(
            current_day_number=day_number,
            current_day_tracks=tracks_in_current_file,
            current_day_plate_lookup_exact=plate_lookup_exact,
            current_day_plate_lookup_by_length=plate_lookup_by_length,
            current_day_start_index=start_index,
            current_day_end_index=end_index,
            current_day_orders_2d=multi_plan_data.get('orders_2d', []),
            current_day_optimization_result=multi_plan_data.get('optimization_result', {}),
            current_day_source_plans=multi_plan_data.get('source_plans', [])
        )
    else:
        await state.update_data(current_day_number=day_number)
    
    await callback.answer()


@router.callback_query(F.data.startswith("generate_breakdown_"))
async def generate_day_breakdown(callback: CallbackQuery, state: FSMContext):
    """
    Генерирует детальную разбивку для выбранного дня.
    
    Простыми словами:
    - Создаёт Excel-файл с подробной информацией о себестоимости
    - Показывает стоимость каждой плиты с учётом переармирования
    - Отправляет только этот файл
    
    НОВАЯ ЛОГИКА:
    - Если from_saved_plan=True: использует данные из мультипланов (current_day_tracks)
    - Если from_saved_plan=False: использует данные из state через _restore_optimization_data
    """
    
    day_number = int(callback.data.split("_")[-1])
    
    await callback.message.answer("📊 Генерирую детальную разбивку...")
    
    data = await state.get_data()
    from_saved_plan = data.get('from_saved_plan', False)
    
    # === НОВАЯ ЛОГИКА: Для сохранённых планов используем данные из мультипланов ===
    if from_saved_plan:
        # Используем данные, сохранённые в process_day_selection
        day_tracks = data.get('current_day_tracks', [])
        
        # FALLBACK: Если данные не загружены, загружаем из всех планов
        if not day_tracks:
            plan_start_date = data.get('plan_start_date')
            if not plan_start_date:
                await callback.message.answer("❌ Нет данных о дорожках для этого дня")
                await callback.answer()
                return
            
            try:
                start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
                selected_date = (start_dt + timedelta(days=day_number - 1)).strftime('%Y-%m-%d')
            except ValueError:
                await callback.message.answer("❌ Ошибка обработки даты")
                await callback.answer()
                return
            
            multi_plan_data = get_tracks_for_date_from_all_plans(selected_date)
            if not multi_plan_data:
                await callback.message.answer("❌ Данные не найдены для этого дня")
                await callback.answer()
                return
            
            day_tracks = multi_plan_data['tracks']
            orders_2d = multi_plan_data.get('orders_2d', [])
            optimization_result = multi_plan_data.get('optimization_result', {})
            
            logger.info(f"[BREAKDOWN] Загружено {len(day_tracks)} дорожек из {multi_plan_data['plans_count']} планов")
        else:
            orders_2d = data.get('current_day_orders_2d', [])
            optimization_result = data.get('current_day_optimization_result', {})
        
        start_index = data.get('current_day_start_index', 0)
        end_index = data.get('current_day_end_index', len(day_tracks))
        tracks_for_this_day = len(day_tracks)
        
        # Восстанавливаем глобальные переменные оптимизации для visualize_plan()
        _restore_optimization_globals(orders_2d, optimization_result)
    else:
        # === СТАРАЯ ЛОГИКА: Для несохранённых планов используем state ===
        restored_data = await _restore_optimization_data(state, day_number)
        
        start_index = restored_data['start_index']
        end_index = restored_data['end_index']
        tracks_for_this_day = restored_data['tracks_for_this_day']
        all_tracks_list = restored_data['all_tracks_list']
        
        # Получаем готовые дорожки для этого дня
        day_tracks = all_tracks_list[start_index:end_index]
    
    try:
        # Генерируем визуализацию с готовыми дорожками (НЕ перегенерируем!)
        result_paths = await asyncio.to_thread(
            visualize_plan,
            output_dir=OUTPUTS_DIR_STR,
            tracks_per_file=None,  # НЕ нужен при existing_tracks
            start_track_index=start_index,
            use_production_pricing=True,
            existing_tracks=day_tracks  # ✅ Используем готовые дорожки из плана!
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
            logger.info(f"[BREAKDOWN] Разбивка для дня {day_number} отправлена: {breakdown_path}")
        else:
            await callback.message.answer("⚠️ Файл разбивки не найден")
    
    except Exception as e:
        logger.exception(f"[BREAKDOWN] Ошибка генерации разбивки для дня {day_number}: {e}")
        await callback.message.answer(
            "❌ Не удалось создать детальную разбивку.\n"
            "Подробности в logs/bot.log."
        )
    
    await callback.answer()


@router.callback_query(F.data.startswith("generate_formovka_"))
async def generate_day_formovka(callback: CallbackQuery, state: FSMContext):
    """
    Генерирует файлы формовки для выбранного дня.
    
    Простыми словами:
    - Создаёт Excel-файлы для каждой дорожки
    - Каждый файл - это шаблон для формовщиков
    - Показывает, какие плиты нужно изготовить
    
    НОВАЯ ЛОГИКА:
    - Если from_saved_plan=True: использует данные из мультипланов
    - Если from_saved_plan=False: использует данные из state
    """
    
    day_number = int(callback.data.split("_")[-1])
    
    await callback.message.answer("📋 Генерирую файлы формовки...")
    
    data = await state.get_data()
    from_saved_plan = data.get('from_saved_plan', False)
    
    # === НОВАЯ ЛОГИКА: Для сохранённых планов используем данные из мультипланов ===
    if from_saved_plan:
        # Используем данные, сохранённые в process_day_selection
        tracks_in_current_file = data.get('current_day_tracks', [])
        
        # FALLBACK: Если данные не загружены, загружаем из всех планов
        if not tracks_in_current_file:
            plan_start_date = data.get('plan_start_date')
            if not plan_start_date:
                await callback.message.answer("❌ Нет данных о дорожках для этого дня")
                await callback.answer()
                return
            
            try:
                start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
                selected_date = (start_dt + timedelta(days=day_number - 1)).strftime('%Y-%m-%d')
            except ValueError:
                await callback.message.answer("❌ Ошибка обработки даты")
                await callback.answer()
                return
            
            multi_plan_data = get_tracks_for_date_from_all_plans(selected_date)
            if not multi_plan_data:
                await callback.message.answer("❌ Данные не найдены для этого дня")
                await callback.answer()
                return
            
            tracks_in_current_file = multi_plan_data['tracks']
            plate_lookup_exact = multi_plan_data['plate_lookup_exact']
            plate_lookup_by_length = multi_plan_data['plate_lookup_by_length']
            
            logger.info(f"[FORMOVKA] Загружено {len(tracks_in_current_file)} дорожек из {multi_plan_data['plans_count']} планов")
        else:
            plate_lookup_exact = data.get('current_day_plate_lookup_exact', {})
            plate_lookup_by_length = data.get('current_day_plate_lookup_by_length', {})
        
        start_index = data.get('current_day_start_index', 0)
        end_index = data.get('current_day_end_index', len(tracks_in_current_file))
    else:
        # === СТАРАЯ ЛОГИКА: Для несохранённых планов используем state ===
        restored_data = await _restore_optimization_data(state, day_number)
        
        start_index = restored_data['start_index']
        end_index = restored_data['end_index']
        all_tracks_list = restored_data['all_tracks_list']
        plate_lookup_exact = restored_data['plate_lookup_exact']
        plate_lookup_by_length = restored_data['plate_lookup_by_length']
        
        # Формируем список плит для формовки
        tracks_in_current_file = all_tracks_list[start_index:end_index]
    
    try:
        
        # Список для хранения данных формовки по всем дорожкам
        formovka_tracks_data = []
        
        # Создаем КОПИЮ lookup для формовки (чтобы не влиять на оригинал в state)
        formovka_lookup_exact = copy.deepcopy(plate_lookup_exact)
        formovka_lookup_by_length = copy.deepcopy(plate_lookup_by_length)
        
        def get_plate_info_smart(length, width):
            """
            Умный поиск информации о плите С УЧЕТОМ КОЛИЧЕСТВА.
            
            FUZZY-ПОИСК: Если точный ключ не найден, ищем с tolerance 0.03м (30мм)
            по длине. Это нужно, потому что оптимизатор может округлять длины.
            """
            TOLERANCE = 0.03  # 30мм tolerance для fuzzy-поиска
            rounded_length = round(length, 2)
            
            # 1. Сначала пробуем точное совпадение
            key = (rounded_length, width)
            entries = formovka_lookup_exact.get(key, [])
            
            for entry in entries:
                if entry.get('qty_remaining', 0) > 0:
                    entry['qty_remaining'] -= 1
                    return entry.copy()
            
            # 2. Если ширина < 1200 (плита с резом), ищем по оригинальной ширине 1200
            if width < 1200:
                key_original = (rounded_length, 1200)
                entries = formovka_lookup_exact.get(key_original, [])
                for entry in entries:
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        return entry.copy()
            
            # 3. Fuzzy-поиск с tolerance по длине в exact lookup
            for lookup_key, entries in formovka_lookup_exact.items():
                key_length, key_width = lookup_key
                # Проверяем ширину (точно или 1200 для split)
                if key_width != width and key_width != 1200:
                    continue
                # Проверяем длину с tolerance
                if abs(key_length - rounded_length) <= TOLERANCE:
                    for entry in entries:
                        if entry.get('qty_remaining', 0) > 0:
                            entry['qty_remaining'] -= 1
                            return entry.copy()
            
            # 4. Fallback: поиск только по длине (точный)
            entries = formovka_lookup_by_length.get(rounded_length, [])
            for entry in entries:
                if entry.get('qty_remaining', 0) > 0:
                    entry['qty_remaining'] -= 1
                    return entry.copy()
            
            # 5. Fuzzy fallback: поиск по длине с tolerance
            for lookup_length, entries in formovka_lookup_by_length.items():
                if abs(lookup_length - rounded_length) <= TOLERANCE:
                    for entry in entries:
                        if entry.get('qty_remaining', 0) > 0:
                            entry['qty_remaining'] -= 1
                            return entry.copy()
            
            return {
                'kp_id': None,
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
                if item is None:
                    continue
                length = item.get('length')
                
                # Определяем ширину в зависимости от режима плиты
                mode = item.get('mode', 'solid')
                if mode == 'transverse' and item.get('width'):
                    width = round(item['width'] * 1000)  # round для корректного округления float
                elif mode == 'split' and item.get('main_w'):
                    width = round(item['main_w'] * 1000)  # round для корректного округления
                else:
                    width = 1200
                
                if not length:
                    continue
                
                plate_info = get_plate_info_smart(length, width)
                # Номер КП: из lookup или из элемента дорожки (item)
                kp_id = plate_info.get('kp_id') or item.get('kp_id')
                
                found = False
                for existing in plates_info:
                    if (round(existing['length'], 2) == round(length, 2) and
                        existing['width'] == width and
                        abs(existing['reinforcement'] - plate_info['reinforcement']) < 0.1 and
                        existing.get('kp_id') == kp_id and
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
                        'kp_id': kp_id,
                        'plate_name': plate_info.get('plate_name', '')
                    })
                
                # Обрабатываем плиты из вторичных резов (остатков)
                secondary_cuts = item.get('secondary_cuts', []) if item else []
                for sec_cut in (secondary_cuts or []):
                    sec_width_m = sec_cut.get('width', 0)
                    if sec_width_m <= 0:
                        continue
                    
                    sec_width = round(sec_width_m * 1000)  # round для корректного округления float
                    sec_length = sec_cut.get('target_length') or length
                    
                    sec_plate_info = get_plate_info_smart(sec_length, sec_width)
                    
                    sec_found = False
                    for existing in plates_info:
                        if (round(existing['length'], 2) == round(sec_length, 2) and
                            existing['width'] == sec_width and
                            existing['kp_date'] == sec_plate_info.get('kp_date', 'неизвестно')):
                            existing['qty'] += 1
                            sec_found = True
                            break
                    
                    if not sec_found:
                        sec_plate_name = sec_plate_info.get('plate_name', '')
                        if not sec_plate_name and sec_cut.get('label'):
                            sec_plate_name = sec_cut['label'].replace('О ', '').strip()
                        
                        plates_info.append({
                            'length': sec_length,
                            'width': sec_width,
                            'qty': 1,
                            'reinforcement': sec_plate_info.get('reinforcement', 0),
                            'kp_date': sec_plate_info.get('kp_date', 'неизвестно'),
                            'customer': sec_plate_info.get('customer', 'неизвестно'),
                            'plate_name': sec_plate_name
                        })
            
            if plates_info:
                # Получаем максимальное армирование дорожки
                max_reinforcement = track.get('max_reinforcement', 0)
                
                # Сортируем плиты по длине
                plates_info.sort(key=lambda x: x['length'], reverse=True)
                
                # Форматируем имена плит если нужно
                for plate in plates_info:
                    if not plate.get('plate_name'):
                        length_dm = int(round(plate['length'] * 10))
                        width_mm = int(plate['width'])
                        
                        load_code = 8
                        if plate.get('reinforcement', 0) > 0:
                            reinforcement = plate['reinforcement']
                            if reinforcement < 8:
                                load_code = 6
                            elif reinforcement < 12:
                                load_code = 8
                            elif reinforcement < 15:
                                load_code = 10
                            else:
                                load_code = 12
                        
                        if width_mm == 1200:
                            width_str = "12"
                        else:
                            width_dm = width_mm / 100.0
                            if abs(width_dm - int(width_dm)) < 0.01:
                                width_str = str(int(width_dm))
                            else:
                                width_str = str(width_dm).replace('.', ',')
                        
                        plate['plate_name'] = f"ПБ {length_dm}-{width_str}-{load_code}п"
                
                # Сохраняем данные для файла формовки
                formovka_tracks_data.append({
                    'track_number': track_number,
                    'max_reinforcement': max_reinforcement,
                    'plates_info': plates_info
                })
        
        # Создаем Excel-файлы формовки
        if formovka_tracks_data:
            template_path = os.path.join(PROJECT_ROOT, "банк знаний", "!КЗ ПБ Шаблон.xlsx")
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            formovka_files = await asyncio.to_thread(
                create_formovka_files_for_tracks,
                formovka_tracks_data,
                OUTPUTS_DIR_STR,
                template_path=template_path,
                date_str=date_str
            )
            
            # Отправляем файлы формовки
            for formovka_file in formovka_files:
                if os.path.exists(formovka_file):
                    track_num = os.path.basename(formovka_file).split('_')[2]
                    await callback.message.answer_document(
                        FSInputFile(formovka_file),
                        caption=f"📋 Формовка (Дорожка {track_num})"
                    )
            
            logger.info(f"[FORMOVKA] Отправлено {len(formovka_files)} файлов формовки для дня {day_number}")
        else:
            await callback.message.answer("⚠️ Нет данных для создания файлов формовки")
    
    except Exception as e:
        logger.exception(f"[FORMOVKA] Ошибка генерации формовки для дня {day_number}: {e}")
        await callback.message.answer(
            "❌ Не удалось создать файлы формовки.\n"
            "Подробности в logs/bot.log."
        )
    
    await callback.answer()


@router.callback_query(F.data.startswith("generate_schema_"))
async def generate_day_schema(callback: CallbackQuery, state: FSMContext):
    """
    Генерирует схему дорожек для выбранного дня.
    
    Простыми словами:
    - Создаёт PDF-файл со схемой раскладки плит
    - Показывает, как плиты расположены на дорожках
    - Отправляет только этот файл (без разбивки и формовки)
    
    НОВАЯ ЛОГИКА:
    - Если from_saved_plan=True: использует данные из мультипланов
    - Если from_saved_plan=False: использует данные из state
    """
    
    day_number = int(callback.data.split("_")[-1])
    
    await callback.message.answer("📐 Генерирую схему дорожек...")
    
    data = await state.get_data()
    from_saved_plan = data.get('from_saved_plan', False)
    
    # === НОВАЯ ЛОГИКА: Для сохранённых планов используем данные из мультипланов ===
    if from_saved_plan:
        # Используем данные, сохранённые в process_day_selection
        day_tracks = data.get('current_day_tracks', [])
        
        # FALLBACK: Если данные не загружены, загружаем из всех планов
        if not day_tracks:
            plan_start_date = data.get('plan_start_date')
            if not plan_start_date:
                await callback.message.answer("❌ Нет данных о дорожках для этого дня")
                await callback.answer()
                return
            
            try:
                start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
                selected_date = (start_dt + timedelta(days=day_number - 1)).strftime('%Y-%m-%d')
            except ValueError:
                await callback.message.answer("❌ Ошибка обработки даты")
                await callback.answer()
                return
            
            multi_plan_data = get_tracks_for_date_from_all_plans(selected_date)
            if not multi_plan_data:
                await callback.message.answer("❌ Данные не найдены для этого дня")
                await callback.answer()
                return
            
            day_tracks = multi_plan_data['tracks']
            orders_2d = multi_plan_data.get('orders_2d', [])
            optimization_result = multi_plan_data.get('optimization_result', {})
            
            logger.info(f"[SCHEMA] Загружено {len(day_tracks)} дорожек из {multi_plan_data['plans_count']} планов")
        else:
            orders_2d = data.get('current_day_orders_2d', [])
            optimization_result = data.get('current_day_optimization_result', {})
        
        start_index = data.get('current_day_start_index', 0)
        end_index = data.get('current_day_end_index', len(day_tracks))
        tracks_for_this_day = len(day_tracks)
        
        # (Глобальные переменные уже восстановлены выше в fallback-логике)
    else:
        # === СТАРАЯ ЛОГИКА: Для несохранённых планов используем state ===
        restored_data = await _restore_optimization_data(state, day_number)
        
        start_index = restored_data['start_index']
        end_index = restored_data['end_index']
        tracks_for_this_day = restored_data['tracks_for_this_day']
        all_tracks_list = restored_data['all_tracks_list']
        
        # Получаем готовые дорожки для этого дня
        day_tracks = all_tracks_list[start_index:end_index]
    
    try:
        # Генерируем визуализацию с готовыми дорожками (НЕ перегенерируем!)
        result_paths = await asyncio.to_thread(
            visualize_plan,
            output_dir=OUTPUTS_DIR_STR,
            tracks_per_file=None,  # НЕ нужен при existing_tracks
            start_track_index=start_index,
            use_production_pricing=True,
            existing_tracks=day_tracks  # ✅ Используем готовые дорожки из плана!
        )
        
        if isinstance(result_paths, tuple) and len(result_paths) >= 2:
            png_path, pdf_schema_path = result_paths
            
            if os.path.exists(pdf_schema_path):
                await callback.message.answer_document(
                    FSInputFile(pdf_schema_path),
                    caption=f"📐 Схема дорожек {start_index + 1}-{end_index} (День {day_number})"
                )
                logger.info(f"[SCHEMA] Схема для дня {day_number} отправлена: {pdf_schema_path}")
            else:
                await callback.message.answer("⚠️ Схема не создана")
        else:
            await callback.message.answer("⚠️ Ошибка генерации схемы")
    
    except Exception as e:
        logger.exception(f"[SCHEMA] Ошибка генерации схемы для дня {day_number}: {e}")
        await callback.message.answer(
            "❌ Не удалось создать схему дорожек.\n"
            "Подробности в logs/bot.log."
        )
    
    await callback.answer()
