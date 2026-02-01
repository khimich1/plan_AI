"""Завершение производственного дня с учётом брака"""
import copy
import logging
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import CallbackQuery
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
import core.config_and_data as cfg

from ..keyboards import main_menu_kb, calendar_days_kb, plates_completion_kb
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import (
    get_active_plan_id, mark_day_completed, get_tracks_for_date_from_all_plans
)

router = Router()


@router.callback_query(F.data.startswith("complete_day_"))
async def start_day_completion(callback: CallbackQuery, state: FSMContext):
    """
    Начало процесса завершения дня.
    Показываем плиты для отметки брака.
    """
    day_number = int(callback.data.split("_")[-1])
    data = await state.get_data()
    
    # НОВОЕ: Проверяем данные из сохранённого плана (current_day_tracks)
    current_day_tracks = data.get('current_day_tracks')
    from_saved_plan = data.get('from_saved_plan', False)
    
    if current_day_tracks:
        # Работаем с сохранённым планом - используем предзагруженные дорожки
        tracks_for_day = current_day_tracks
        plate_lookup_exact = data.get('current_day_plate_lookup_exact', {})
        plate_lookup_by_length = data.get('current_day_plate_lookup_by_length', {})
        start_index = data.get('current_day_start_index', 0)
        source_plans = data.get('current_day_source_plans', [])
        
        logger.info(f"[COMPLETION] День {day_number}: используем предзагруженные дорожки ({len(tracks_for_day)} шт)")
    elif from_saved_plan:
        # МУЛЬТИПЛАН: Работаем с сохранённым планом, но данные не загружены
        # Загружаем дорожки из ВСЕХ планов на эту дату
        plan_start_date = data.get('plan_start_date')
        
        if not plan_start_date:
            await callback.message.answer(
                "❌ Не удалось определить дату начала плана.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            await callback.answer()
            return
        
        # Вычисляем дату выбранного дня
        try:
            start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
            selected_date = (start_dt + timedelta(days=day_number - 1)).strftime('%Y-%m-%d')
        except ValueError as e:
            logger.error(f"[COMPLETION] Ошибка парсинга даты: {e}")
            await callback.message.answer(
                "❌ Ошибка обработки даты.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            await callback.answer()
            return
        
        # Загружаем дорожки из всех планов на эту дату
        logger.info(f"[COMPLETION] День {day_number} ({selected_date}): загружаем из всех планов...")
        multi_plan_data = get_tracks_for_date_from_all_plans(selected_date)
        
        if not multi_plan_data or not multi_plan_data.get('tracks'):
            await callback.message.answer(
                f"❌ Дата {datetime.strptime(selected_date, '%Y-%m-%d').strftime('%d.%m.%Y')} "
                f"не найдена ни в одном сохранённом плане.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            await callback.answer()
            return
        
        tracks_for_day = multi_plan_data['tracks']
        plate_lookup_exact = multi_plan_data['plate_lookup_exact']
        plate_lookup_by_length = multi_plan_data['plate_lookup_by_length']
        start_index = 0
        source_plans = multi_plan_data.get('source_plans', [])
        
        logger.info(f"[COMPLETION] День {day_number}: загружено {len(tracks_for_day)} дорожек из {multi_plan_data['plans_count']} планов")
        
        # Сохраняем source_plans для последующей отметки завершения во всех планах
        await state.update_data(current_day_source_plans=source_plans)
    else:
        # СТАРАЯ ЛОГИКА: Работаем с новым планом - извлекаем из all_tracks_list
        tracks_count = data.get('tracks_count', 1)
        all_tracks_list = data.get('all_tracks_list', [])
        plate_lookup_exact = data.get('plate_lookup_exact', {})
        plate_lookup_by_length = data.get('plate_lookup_by_length', {})
        source_plans = []
        
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
        
        logger.info(f"[COMPLETION] День {day_number}: используем данные из нового плана ({len(tracks_for_day)} дорожек)")
    
    # Создаем КОПИЮ lookup для завершения дня (чтобы не влиять на оригинал в state)
    completion_lookup_exact = copy.deepcopy(plate_lookup_exact)
    completion_lookup_by_length = copy.deepcopy(plate_lookup_by_length)
    
    # Функция для получения информации о плите (с списанием из lookup)
    def get_plate_info_smart(length, width, expected_kp_id=None):
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
        
        # ✅ ЛОГИРУЕМ ПОИСК ПЛИТЫ
        logger.debug(f"[TRACE] Ищем плиту: длина={length:.2f}м ({rounded_length:.2f}м), ширина={width}мм")
        
        # 1. Сначала пробуем точное совпадение
        key = (rounded_length, width)
        entries = completion_lookup_exact.get(key, [])
        
        if entries:
            logger.debug(f"[TRACE]   Найдено {len(entries)} записей для ключа {key}")
        
        for entry in entries:
            if expected_kp_id and entry.get('kp_id') != expected_kp_id:
                continue
            if entry.get('qty_remaining', 0) > 0:
                entry['qty_remaining'] -= 1
                return entry.copy()
        
        # 2. Fuzzy-поиск с tolerance по длине И ширине в exact lookup
        # ИСПРАВЛЕНИЕ: Убран некорректный поиск по ширине 1200 для плит с резом,
        # так как это вызывало проблемы для вторичных резов (плиты 460мм искались как 1200мм)
        logger.debug(f"[TRACE]   Пробуем fuzzy-поиск (tolerance={TOLERANCE}м)")
        for lookup_key, entries in completion_lookup_exact.items():
            key_length, key_width = lookup_key
            # Проверяем ширину с tolerance 20мм (для небольших отклонений типа 460 vs 459)
            if abs(key_width - width) > 20:
                continue
            # Проверяем длину с tolerance
            if abs(key_length - rounded_length) <= TOLERANCE:
                for entry in entries:
                    if expected_kp_id and entry.get('kp_id') != expected_kp_id:
                        continue
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        logger.debug(f"[TRACE]   ✓ Найдено FUZZY совпадение: ключ={lookup_key}, КП #{entry.get('kp_id', '?')}")
                        return entry.copy()
        
        # 4. Fallback: поиск только по длине (точный)
        entries = completion_lookup_by_length.get(rounded_length, [])
        for entry in entries:
            if expected_kp_id and entry.get('kp_id') != expected_kp_id:
                continue
            if entry.get('qty_remaining', 0) > 0:
                entry['qty_remaining'] -= 1
                return entry.copy()
        
        # 5. Fuzzy fallback: поиск по длине с tolerance
        for lookup_length, entries in completion_lookup_by_length.items():
            if abs(lookup_length - rounded_length) <= TOLERANCE:
                for entry in entries:
                    if expected_kp_id and entry.get('kp_id') != expected_kp_id:
                        continue
                    if entry.get('qty_remaining', 0) > 0:
                        entry['qty_remaining'] -= 1
                        return entry.copy()
        
        # ✅ ЛОГИРУЕМ ОШИБКУ ПОИСКА
        logger.warning(f"[TRACE]   ❌ НЕ НАЙДЕНО совпадение для: длина={length:.2f}м, ширина={width}мм")
        logger.warning(f"[TRACE]   Доступные ключи в lookup_exact: {list(completion_lookup_exact.keys())[:10]}")
        
        return {
            'kp_date': 'неизвестно',
            'customer': 'неизвестно',
            'plate_name': '',
            'kp_id': None
        }
    
    # ✅ НОВОЕ: Логируем плиты ДО обработки для списания
    logger.info(f"[TRACE] ===== ШАГ 7: ПЛИТЫ ПЕРЕД СПИСАНИЕМ (День {day_number}) =====")
    logger.info(f"[TRACE] Дорожек для этого дня: {len(tracks_for_day)}")
    
    total_items_before = 0
    for track_idx, track in enumerate(tracks_for_day):
        items_count = len(track.get('items', []))
        total_items_before += items_count
        logger.info(f"[TRACE]   Дорожка #{start_index + track_idx + 1}: {items_count} плит")
    
    logger.info(f"[TRACE] Всего плит в дорожках: {total_items_before}")
    
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
                width = round(item['width'] * 1000)  # round для корректного округления float
            elif mode == 'split' and item.get('main_w'):
                width = round(item['main_w'] * 1000)  # round для корректного округления
            else:
                # solid mode - берём width из поля (ИСПРАВЛЕНИЕ: учитываем реальную ширину)
                width = round(item.get('width', 1.2) * 1000)  # round для корректного округления
            
            if not length:
                continue
            
            item_kp_id = item.get('kp_id')
            item_plate_name = item.get('plate_name')
            item_load_code = cfg.normalize_load_code(item.get('load_code', 8))

            plate_info = get_plate_info_smart(length, width, expected_kp_id=item_kp_id)
            plate_name = plate_info.get('plate_name', '') or (item_plate_name or '')
            kp_id = plate_info.get('kp_id') or item_kp_id
            
            # Если нет имени плиты — формируем его
            if not plate_name and kp_id:
                # Пытаемся взять точное имя из БД по kp_id+размерам
                try:
                    conn = None
                    db_path = str(PROJECT_ROOT / "plita.db")
                    conn = sqlite3.connect(db_path)
                    cur = conn.cursor()
                    cur.execute('''
                        SELECT plate_name, load_class
                        FROM kp_plates
                        WHERE kp_id = ?
                          AND status = 'в плане'
                          AND ABS(length_m - ?) < 0.02
                          AND ABS(width_m - ?) < 0.01
                        LIMIT 1
                    ''', (kp_id, length, width / 1000.0))
                    row = cur.fetchone()
                    if row:
                        plate_name = row[0]
                        item_load_code = cfg.normalize_load_code((row[1] or 800) / 100)
                finally:
                    try:
                        conn.close()
                    except Exception:
                        pass

            if not plate_name:
                plate_name = cfg.make_plate_name(length, width / 1000.0, load_code=item_load_code)
            
            # Получаем дату и заказчика для группировки
            kp_date = plate_info.get('kp_date', 'неизвестно') or item.get('kp_date', 'неизвестно')
            customer = plate_info.get('customer', 'неизвестно') or item.get('customer', 'неизвестно')
            
            # Ищем такую же плиту в списке текущей дорожки
            # Группируем по: plate_name + kp_id + kp_date + customer + width
            found = False
            width_m = width / 1000.0
            for existing in track_plates:
                if (existing['plate_name'] == plate_name and 
                    existing['kp_id'] == kp_id and
                    existing['kp_date'] == kp_date and
                    existing['customer'] == customer and
                    abs(existing['width_m'] - width_m) < 0.01 and
                    not existing.get('is_secondary', False)):  # Только с другими основными!
                    existing['qty'] += 1
                    found = True
                    break
            
            if not found:
                track_plates.append({
                    'plate_name': plate_name,
                    'length_m': length,
                    'width_m': width_m,
                    'load_class': int(cfg.normalize_load_code(item_load_code) * 100),
                    'qty': 1,
                    'kp_id': kp_id,
                    'kp_date': kp_date,
                    'customer': customer,
                    'is_secondary': False  # Флаг: это основная плита
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
                
                sec_plate_info = get_plate_info_smart(sec_length, sec_width, expected_kp_id=item_kp_id)
                sec_plate_name = sec_plate_info.get('plate_name', '') or (sec_cut.get('label', '') or '').replace('О ', '').strip()
                sec_kp_id = sec_plate_info.get('kp_id') or item_kp_id
                
                # Если нет имени плиты — берём из label
                if not sec_plate_name:
                    sec_plate_name = cfg.make_plate_name(sec_length, sec_width / 1000.0, load_code=item_load_code)
                
                sec_kp_date = sec_plate_info.get('kp_date', 'неизвестно') or item.get('kp_date', 'неизвестно')
                sec_customer = sec_plate_info.get('customer', 'неизвестно') or item.get('customer', 'неизвестно')
                sec_width_m = sec_width / 1000.0
                
                # Ищем такую же плиту в списке (только среди вторичных!)
                sec_found = False
                for existing in track_plates:
                    if (existing['plate_name'] == sec_plate_name and 
                        existing['kp_id'] == sec_kp_id and
                        existing['kp_date'] == sec_kp_date and
                        existing['customer'] == sec_customer and
                        abs(existing['width_m'] - sec_width_m) < 0.01 and
                        existing.get('is_secondary', False) == True):  # Только с другими вторичными!
                        existing['qty'] += 1
                        sec_found = True
                        break
                
                if not sec_found:
                    track_plates.append({
                        'plate_name': sec_plate_name,
                        'length_m': sec_length,
                        'width_m': sec_width_m,
                        'load_class': int(cfg.normalize_load_code(item_load_code) * 100),
                        'qty': 1,
                        'kp_id': sec_kp_id,
                        'kp_date': sec_kp_date,
                        'customer': sec_customer,
                        'is_secondary': True  # Флаг: это вторичный рез
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
        plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
        completed_days = data.get('completed_days', [])
        days_info = data.get('days_info', {})
        from_saved_plan = data.get('from_saved_plan', False)
        
        await callback.message.answer(
            f"❌ Не удалось найти плиты для Дня {day_number}.",
            reply_markup=calendar_days_kb(
                data.get('total_days', 1), 
                plan_start_date, 
                completed_days,
                days_info,
                show_save_button=not from_saved_plan  # Скрываем кнопку если план уже сохранён
            )
        )
        await callback.answer()
        return
    
    # Подсчитываем общее количество позиций
    total_positions = sum(len(track['plates']) for track in day_plates_by_track)
    
    # ✅ НОВОЕ: Логируем результаты формирования списка для списания
    logger.info(f"[TRACE] ===== ШАГ 8: ПЛИТЫ ДЛЯ СПИСАНИЯ (day_plates_by_track) =====")
    logger.info(f"[TRACE] Дорожек: {len(day_plates_by_track)}")
    logger.info(f"[TRACE] Позиций (уникальных плит): {total_positions}")
    logger.info(f"[TRACE] Всего плит (с количеством): {total_qty}")
    
    for track_data in day_plates_by_track:
        track_num = track_data['track_number']
        plates_count = len(track_data['plates'])
        qty_sum = sum(p['qty'] for p in track_data['plates'])
        logger.info(f"[TRACE]   Дорожка #{track_num}: {plates_count} позиций, {qty_sum} шт")
        
        for plate in track_data['plates']:
            is_sec = " [ВТОРИЧНЫЙ]" if plate.get('is_secondary') else ""
            logger.info(f"[TRACE]     - {plate['plate_name']} × {plate['qty']} (ширина={plate['width_m']*1000:.0f}мм, КП #{plate.get('kp_id', '?')}){is_sec}")
    
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
    for plate in plates_without_kp:
        length_m = plate.get('length_m', 0)
        plate_name = plate.get('plate_name', '')
        
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        
        # Ищем по длине среди КП "в работе" и плит "в плане"
        cur.execute('''
            SELECT p.kp_id, p.plate_name 
            FROM kp_plates p
            JOIN kp_meta m ON p.kp_id = m.kp_id
            WHERE ABS(p.length_m - ?) < 0.02 
              AND p.qty > 0
              AND m.status = 'в работе'
              AND p.status = 'в плане'
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
    
    # ========== ВОЗВРАТ БРАКОВАННЫХ ПЛИТ В ПРОИЗВОДСТВО ==========
    # Бракованные плиты возвращаются в статус 'в производстве',
    # чтобы попасть в следующее планирование
    rejected_returned = 0
    for plate in rejected_plates:
        kp_id = plate.get('kp_id')
        plate_name = plate.get('plate_name')
        qty = plate.get('qty', 1)
        
        if kp_id and plate_name and qty > 0:
            success = kp_db.return_plates_to_production(
                kp_id=kp_id,
                plate_name=plate_name,
                qty=qty,
                db_path=db_path
            )
            if success:
                rejected_returned += 1
                logger.info(f"[COMPLETION] Брак: {plate_name} x{qty} возвращена в производство (КП #{kp_id})")
    
    if rejected_returned > 0:
        logger.info(f"[COMPLETION] Всего возвращено в производство: {rejected_returned} позиций (брак)")
    # ========== КОНЕЦ ВОЗВРАТА БРАКА ==========
    
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
    
    # Показываем календарь с обновлённой галочкой
    total_days = data.get('total_days', 1)
    start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    days_info = data.get('days_info', {})
    completed_days = data.get('completed_days', [])
    
    # Обновляем days_info - помечаем текущий день как завершённый
    completing_day = day_number
    current_day_date = None
    try:
        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
        current_day_date = (start_dt + timedelta(days=completing_day - 1)).strftime('%Y-%m-%d')
    except:
        pass
    
    if current_day_date and current_day_date in days_info:
        days_info[current_day_date]['completed'] = True
    
    # ВАЖНО: Сохраняем изменения в файл плана!
    active_plan_id = data.get('active_plan_id') or get_active_plan_id()
    source_plans = data.get('current_day_source_plans', [])
    
    # Если есть список планов-источников (мультиплан), отмечаем день во всех планах
    if source_plans:
        for plan_id in source_plans:
            if mark_day_completed(plan_id, current_day_date):
                logger.info(f"День {current_day_date} отмечен как завершённый в плане {plan_id}")
            else:
                logger.warning(f"Не удалось отметить день {current_day_date} как завершённый в плане {plan_id}")
    elif active_plan_id and current_day_date:
        # Если работаем с одним активным планом
        if mark_day_completed(active_plan_id, current_day_date):
            logger.info(f"День {current_day_date} отмечен как завершённый в плане {active_plan_id}")
        else:
            logger.warning(f"Не удалось отметить день {current_day_date} как завершённый")
    
    # Обновляем completed_days в state
    if day_number not in completed_days:
        completed_days.append(day_number)
        completed_days.sort()
        await state.update_data(completed_days=completed_days)
    
    # Проверяем, работаем ли мы с сохранённым планом
    from_saved_plan = data.get('from_saved_plan', False)
    
    await callback.message.answer(report)
    await callback.message.answer(
        "📅 Календарь производства обновлён!\n"
        "Выберите следующий день:",
        reply_markup=calendar_days_kb(
            total_days, 
            start_date, 
            completed_days,
            days_info,
            show_save_button=not from_saved_plan  # Скрываем кнопку если план уже сохранён
        )
    )
    await state.set_state(ProductionStates.viewing_calendar)
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
    
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = data.get('days_info', {})
    
    # Проверяем, работаем ли мы с сохранённым планом
    from_saved_plan = data.get('from_saved_plan', False)
    
    await callback.message.answer(
        "❌ Завершение дня отменено.\n\n"
        "Выберите день для просмотра:",
        reply_markup=calendar_days_kb(
            total_days, 
            plan_start_date, 
            completed_days,
            days_info,
            show_save_button=not from_saved_plan  # Скрываем кнопку если план уже сохранён
        )
    )
    await state.set_state(ProductionStates.viewing_calendar)
    await callback.answer()
