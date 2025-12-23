"""Обработчики планирования производства плит"""
import asyncio
import os
from datetime import datetime
from pathlib import Path
from collections import defaultdict
import math

from aiogram import Router, F
from aiogram.types import Message, FSInputFile
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

from ..keyboards import main_menu_kb
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
    """Получаем дату (полную или число месяца) и запускаем планирование"""
    user_input = message.text.strip()
    
    # === ПАРСИНГ ДАТЫ: поддерживаем разные форматы ===
    target_date = None
    
    # Формат 1: Полная дата "ДД.ММ.ГГГГ" (например: "01.02.2026")
    try:
        target_date = datetime.strptime(user_input, '%d.%m.%Y')
        date_description = target_date.strftime('%d.%m.%Y')
    except ValueError:
        pass
    
    # Формат 2: Полная дата "ГГГГ-ММ-ДД" (например: "2026-02-01")
    if not target_date:
        try:
            target_date = datetime.strptime(user_input, '%Y-%m-%d')
            date_description = target_date.strftime('%d.%m.%Y')
        except ValueError:
            pass
    
    # Формат 3: Только число месяца (например: "25" = 25 число текущего месяца)
    if not target_date:
        try:
            date_number = int(user_input)
            
            if date_number < 1 or date_number > 31:
                await message.answer(
                    "❌ Число должно быть от 1 до 31.\n"
                    "Попробуйте снова:"
                )
                return
            
            # Берём текущий месяц/год
            now = datetime.now()
            target_date = datetime(now.year, now.month, date_number)
            date_description = f"{date_number} {target_date.strftime('%B %Y')}"
        except ValueError:
            pass
    
    # Если не удалось распознать формат
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
    
    # Получаем данные из состояния
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    
    await message.answer(
        f"✅ Параметры планирования:\n"
        f"• Дорожек: {tracks_count}\n"
        f"• Плиты со сроком до: {date_description}\n\n"
        f"📌 Логика выбора:\n"
        f"1️⃣ Сначала беру плиты с более ранней датой\n"
        f"2️⃣ Внутри даты — с меньшим армированием\n"
        f"3️⃣ Если не хватает — беру с той же датой, но большим армированием\n"
        f"4️⃣ Когда день закончится — перехожу к следующему дню\n\n"
        f"⏳ Загружаю плиты из базы данных..."
    )
    
    try:
        # === ШАГ 1: ВЫБОРКА ПЛИТ ИЗ БД ===
        db_path = PROJECT_ROOT / "plita.db"
        pb_db_path = BOT_DIR / "pb.db"
        
        # Получаем все КП в работе с дедлайном <= target_date
        import sqlite3
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        
        # Получаем КП со сроком до указанной даты
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
            
            # Парсим дату (формат "ДД.ММ.ГГГГ")
            try:
                exec_date = datetime.strptime(exec_terms, '%d.%m.%Y')
                
                # Сравниваем даты: берём только те КП, где срок <= target_date
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
        
        # Сортируем КП по дате (раньше → позже)
        kp_list.sort(key=lambda x: x['date'])
        
        await message.answer(
            f"✅ Найдено КП в работе: {len(kp_list)}\n"
            f"Загружаю плиты из этих КП..."
        )
        
        # === ШАГ 2: СОБИРАЕМ ВСЕ ПЛИТЫ ИЗ ЭТИХ КП ===
        # Группируем: сначала по дате КП, потом по армированию
        plates_by_date_and_reinforcement = defaultdict(lambda: defaultdict(list))
        
        for kp_info in kp_list:
            kp_id = kp_info['kp_id']
            kp_date = kp_info['date']
            
            # Получаем плиты из этого КП
            cur.execute("""
                SELECT plate_name, length_m, width_m, load_class, qty
                FROM kp_plates
                WHERE kp_id = ?
            """, (kp_id,))
            
            for row in cur.fetchall():
                plate_name, length_m, width_m, load_class, qty = row
                
                # Получаем армирование из БД pb.db
                load_code = load_class // 100  # 800 → 8
                reinforcement_value = get_reinforcement(
                    length_m=length_m,
                    load_code=load_code,
                    source='series',
                    db_path=pb_db_path,
                    allow_fallback=True
                )
                
                # Если армирование не найдено, используем fallback
                if reinforcement_value is None:
                    reinforcement_value = 999.0  # Большое число для сортировки в конец
                
                # Группируем: сначала по дате КП, потом по армированию
                plates_by_date_and_reinforcement[kp_date][reinforcement_value].append({
                    'plate_name': plate_name,
                    'length': length_m,
                    'width': int(width_m * 1000),  # метры → мм
                    'load_code': load_code,
                    'qty': qty,
                    'reinforcement': reinforcement_value,
                    'kp_id': kp_id,
                    'kp_date': kp_date.strftime('%d.%m.%Y'),
                    'customer': kp_info['customer']  # Сохраняем заказчика
                })
        
        # Создаём словарь для быстрого поиска информации о КП по kp_id
        kp_info_dict = {kp['kp_id']: {
            'kp_date': kp['date'].strftime('%d.%m.%Y'),
            'customer': kp['customer']
        } for kp in kp_list}
        
        # Создаём словарь соответствия плит и КП: (length, width) → (kp_id, kp_date, customer, plate_name)
        plate_to_kp_info = {}  # {(length, width): {'kp_id': ..., 'kp_date': ..., 'customer': ..., 'plate_name': ...}}
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
                        'plate_name': plate_name  # Сохраняем название плиты
                    }
        
        conn.close()
        
        # === ШАГ 3: ФОРМИРУЕМ СПИСОК ПЛИТ ДЛЯ ЗАГРУЗКИ ===
        # ЛОГИКА (улучшенная):
        # 1. Берём плиты с самой ранней датой + меньшим армированием
        # 2. Если не хватает на линию → берём с той же датой, но большим армированием
        # 3. Когда все плиты на день закончились → переходим к следующему дню, опять с меньшего армирования
        # 4. Берём ВСЕ плиты, а фильтрацию делаем ПОСЛЕ оптимизации
        
        selected_plates = []
        all_plates_by_order = []  # Все плиты в порядке приоритета
        
        # Создаём сводку для пользователя
        selection_log = []
        
        # Сортируем даты КП по возрастанию (от более ранних к поздним)
        sorted_dates = sorted(plates_by_date_and_reinforcement.keys())
        
        # Собираем ВСЕ плиты в порядке приоритета (дата → армирование)
        for kp_date in sorted_dates:
            reinforcement_dict = plates_by_date_and_reinforcement[kp_date]
            
            # Сортируем армирование по возрастанию (от меньшего к большему)
            sorted_reinforcements = sorted(reinforcement_dict.keys())
            
            date_str = kp_date.strftime('%d.%m.%Y')
            date_plates_count = 0
            
            for reinforcement in sorted_reinforcements:
                plates = reinforcement_dict[reinforcement]
                
                for plate_data in plates:
                    all_plates_by_order.append(plate_data)
                    selected_plates.append(plate_data)
                    date_plates_count += plate_data['qty']
                    
                    # Логируем взятую плиту
                    selection_log.append(
                        f"  📦 {plate_data['plate_name']} × {plate_data['qty']} шт "
                        f"(армир. {reinforcement:.1f}, срок {date_str})"
                    )
            
            # Сообщаем о количестве плит с этой даты
            if date_plates_count > 0:
                selection_log.append(
                    f"✅ Взято плит со сроком {date_str}: {date_plates_count} шт\n"
                )
        
        print(f"[ПЛАНИРОВАНИЕ] Собрано {len(selected_plates)} типов плит для оптимизации")
        
        if not selected_plates:
            await message.answer(
                "❌ Не найдено плит для планирования.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        # Показываем детальную информацию о выборе плит
        total_selected_qty = sum(p['qty'] for p in selected_plates)
        total_selected_length = sum(p['length'] * p['qty'] for p in selected_plates)
        
        summary_text = (
            f"✅ Отобрано плит: {total_selected_qty} шт ({len(selected_plates)} типов)\n"
            f"📏 Общая длина: {total_selected_length:.1f}м\n"
            f"🎯 Целевое количество дорожек: {tracks_count}\n\n"
            f"📋 Детали выбора:\n"
        )
        
        # Добавляем лог выбора (до 4000 символов - лимит Telegram)
        log_text = "\n".join(selection_log)
        if len(summary_text + log_text) > 3900:
            # Обрезаем лог, если слишком длинный
            log_text = log_text[:3700] + "\n...\n(список обрезан)"
        
        await message.answer(summary_text + log_text)
        await message.answer("⏳ Запускаю оптимизацию раскроя...")
        
        # === ШАГ 4: ЗАГРУЖАЕМ В ОПТИМИЗАТОР ===
        # Конвертируем в формат для оптимизатора (включая kp_date, customer и plate_name для статистики)
        orders_2d = []
        for plate_data in selected_plates:
            orders_2d.append({
                'length': plate_data['length'],
                'width': plate_data['width'],
                'qty': plate_data['qty'],
                'load_code': plate_data['load_code'],
                'reinforcement': plate_data['reinforcement'],
                'kp_date': plate_data.get('kp_date', 'неизвестно'),  # Для статистики
                'customer': plate_data.get('customer', 'неизвестно'),  # Заказчик
                'plate_name': plate_data.get('plate_name', '')  # Название плиты
            })
        
        # === ШАГ 4: ЗАПУСКАЕМ ОПТИМИЗАЦИЮ (ОДИН РАЗ) ===
        optimization_result = await asyncio.to_thread(
            optimize_with_cascading_longitudinal_cuts,
            orders_2d=orders_2d
        )
        
        if not optimization_result or optimization_result.get('total_plates', 0) == 0:
            await message.answer(
                "❌ Оптимизация не дала результатов. Проверьте данные плит.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        total_plates_optimized = optimization_result.get('total_plates', 0)
        await message.answer(f"✅ Оптимизация завершена! Исходных плит: {total_plates_optimized}")
        
        # === ШАГ 5: СОХРАНЯЕМ ВСЕ ДАННЫЕ ДЛЯ ВИЗУАЛИЗАЦИИ ===
        # Сохраняем ПОЛНЫЙ результат оптимизации (без фильтрации)
        optimization.OPT_CASCADING_PLAN = optimization_result
        
        # Для совместимости с визуализацией
        all_loads = set(p['load_code'] for p in orders_2d)
        optimization_result['loads_in_group'] = sorted(all_loads)
        optimization.OPT_CASCADING_PLAN_BY_LOAD = {'all': optimization_result}
        optimization.LOAD_TO_REINFORCEMENT_MAP = {
            load_code: ['all'] for load_code in all_loads
        }
        
        # Обновляем cfg для визуализации (ВСЕ плиты)
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
        
        # === ШАГ 6: ВИЗУАЛИЗАЦИЯ С РАЗБИВКОЙ НА ФАЙЛЫ ===
        await message.answer("⏳ Генерирую схемы дорожек...")
        
        # Сначала получаем последовательность, чтобы узнать количество дорожек
        from viz_modules.layout_sequence import build_layout_sequence
        seq = build_layout_sequence()
        
        # Подсчитываем общее количество дорожек (логика из visualize_plan)
        MAX_TRACK_LENGTH = 101.0
        all_tracks_list = []
        
        if isinstance(seq, list) and seq and isinstance(seq[0], dict) and 'load_code' in seq[0]:
            # Группировка по нагрузкам
            for group in seq:
                load_code = group['load_code']
                items = group['sequence']
                group_label = group.get('label', f'Нагрузка {load_code}п')
                
                current_track = []
                current_track_length = 0.0
                
                for i, item in enumerate(items):
                    item_length = item['length']
                    is_solid = (item.get('mode') == 'solid')
                    
                    # Проверяем: добавление плиты превысит лимит?
                    will_exceed = (current_track_length + item_length > MAX_TRACK_LENGTH and current_track)
                    
                    if will_exceed:
                        # ПРАВИЛО: Не превышаем 101м НИКОГДА!
                        # Закрываем дорожку (даже если она будет короткой)
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
            # Старый формат
            current_track = []
            current_track_length = 0.0
            
            for i, item in enumerate(seq):
                item_length = item['length']
                is_solid = (item.get('mode') == 'solid')
                
                # Проверяем: добавление плиты превысит лимит?
                will_exceed = (current_track_length + item_length > MAX_TRACK_LENGTH and current_track)
                
                if will_exceed:
                    # ПРАВИЛО: Не превышаем 101м НИКОГДА!
                    # Закрываем дорожку (даже если она будет короткой)
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
        await message.answer(
            f"📊 Всего дорожек получилось: {total_tracks_count}\n"
            f"📋 Требуется дорожек на сегодня: {tracks_count}"
        )
        
        # Разбиваем дорожки на файлы
        file_number = 1
        start_index = 0
        
        while start_index < total_tracks_count:
            tracks_in_file = min(tracks_count, total_tracks_count - start_index)
            
            await message.answer(
                f"📄 Создаю файл {file_number} (дорожки {start_index + 1}-{start_index + tracks_in_file})..."
            )
            
            # Вызываем визуализацию для этого диапазона дорожек
            result_paths = await asyncio.to_thread(
                visualize_plan,
                OUTPUTS_DIR_STR,
                tracks_in_file,
                start_index
            )
            
            if isinstance(result_paths, tuple) and len(result_paths) >= 2:
                png_path, pdf_schema_path = result_paths
                
                if os.path.exists(pdf_schema_path):
                    first_track = start_index + 1
                    last_track = start_index + tracks_in_file
                    await message.answer_document(
                        FSInputFile(pdf_schema_path),
                        caption=f"📐 Схема дорожек {first_track}-{last_track} (файл {file_number})"
                    )
                
                # Ищем детальную разбивку
                # ВАЖНО: не используем точный timestamp, т.к. он может не совпадать (микросекунды)
                # Ищем по маске: начало имени + номера дорожек
                first_track = start_index + 1
                last_track = start_index + tracks_in_file
                
                if tracks_in_file == 1:
                    breakdown_pattern_prefix = f'Детальная_разбивка_Дорожка_{first_track}_'
                else:
                    breakdown_pattern_prefix = f'Детальная_разбивка_Дорожки_{first_track}-{last_track}_'
                
                # Ищем все файлы, начинающиеся с этого префикса
                breakdown_path = None
                try:
                    for filename in os.listdir(OUTPUTS_DIR_STR):
                        if filename.startswith(breakdown_pattern_prefix) and filename.endswith('.xlsx'):
                            # Берём самый свежий файл (по времени создания)
                            candidate_path = os.path.join(OUTPUTS_DIR_STR, filename)
                            if breakdown_path is None or os.path.getctime(candidate_path) > os.path.getctime(breakdown_path):
                                breakdown_path = candidate_path
                except Exception as e:
                    print(f"[DEBUG] Ошибка поиска файла разбивки: {e}")
                
                if breakdown_path and os.path.exists(breakdown_path):
                    await message.answer_document(
                        FSInputFile(breakdown_path),
                        caption=f"📊 Детальная разбивка (файл {file_number})"
                    )
                
                # === ФОРМИРУЕМ СПИСОК ПЛИТ ПО ДОРОЖКАМ ===
                # Собираем информацию о плитах из дорожек этого файла
                tracks_in_current_file = all_tracks_list[start_index:start_index + tracks_in_file]
                
                # Функция для поиска исходных данных плиты (с приоритетом поиска из БД)
                def find_plate_info(length, width):
                    """Находит информацию о плите из orders_2d или напрямую из БД"""
                    # Сначала пробуем найти по точным размерам в plate_to_kp_info
                    key = (round(length, 2), width)
                    if key in plate_to_kp_info:
                        kp_data = plate_to_kp_info[key]
                        # Ищем армирование и название в orders_2d
                        reinforcement = 0
                        plate_name = kp_data.get('plate_name', '')
                        for order in orders_2d:
                            if (abs(order['length'] - length) < 0.1 and 
                                abs(order['width'] - width) < 50):
                                reinforcement = order.get('reinforcement', 0)
                                # Если есть название в orders_2d, используем его
                                if order.get('plate_name'):
                                    plate_name = order.get('plate_name')
                                break
                        
                        return {
                            'reinforcement': reinforcement,
                            'kp_date': kp_data['kp_date'],
                            'customer': kp_data['customer'],
                            'plate_name': plate_name
                        }
                    
                    # НОВОЕ: Ищем БЛИЖАЙШУЮ плиту по длине (на случай резов)
                    # Сначала ищем среди плит той же ширины в БД
                    best_plate_match = None
                    min_length_diff = float('inf')
                    
                    for (plate_len, plate_width), kp_data in plate_to_kp_info.items():
                        if plate_width == width:
                            length_diff = abs(plate_len - length)
                            # Допуск: ±1 метр (на случай поперечного реза)
                            if length_diff < 1.0 and length_diff < min_length_diff:
                                min_length_diff = length_diff
                                best_plate_match = kp_data
                    
                    if best_plate_match:
                        # Нашли похожую плиту в БД
                        reinforcement = 0
                        for order in orders_2d:
                            if (abs(order['length'] - length) < 1.0 and 
                                abs(order['width'] - width) < 50):
                                reinforcement = order.get('reinforcement', 0)
                                break
                        
                        return {
                            'reinforcement': reinforcement,
                            'kp_date': best_plate_match['kp_date'],
                            'customer': best_plate_match['customer'],
                            'plate_name': best_plate_match.get('plate_name', '')
                        }
                    
                    # Fallback: ищем в orders_2d с допуском
                    best_match = None
                    best_diff = float('inf')
                    
                    for order in orders_2d:
                        len_diff = abs(order['length'] - length)
                        width_diff = abs(order['width'] - width)
                        
                        if len_diff < 1.0 and width_diff < 50:
                            total_diff = len_diff + width_diff / 1000.0
                            if total_diff < best_diff:
                                best_diff = total_diff
                                best_match = {
                                    'reinforcement': order.get('reinforcement', 0),
                                    'kp_date': order.get('kp_date', 'неизвестно'),
                                    'customer': order.get('customer', 'неизвестно'),
                                    'plate_name': order.get('plate_name', '')
                                }
                    
                    return best_match
                
                # Формируем список плит по дорожкам
                for track_idx_in_file, track in enumerate(tracks_in_current_file):
                    track_number = start_index + track_idx_in_file + 1
                    track_items = track.get('items', [])
                    
                    if not track_items:
                        continue
                    
                    # Группируем плиты по уникальным комбинациям (длина, ширина, армирование, дата, заказчик)
                    plates_info = []
                    for item in track_items:
                        length = item.get('length')
                        width = item.get('width', 1200)
                        
                        if not length:
                            continue
                        
                        # Ищем информацию о плите
                        # ВАЖНО: Сначала ищем в orders_2d (там точная информация)
                        plate_info = None
                        
                        # Ищем в orders_2d с допуском по длине
                        for order in orders_2d:
                            if (abs(order['length'] - length) < 0.1 and 
                                abs(order['width'] - width) < 50):
                                plate_info = {
                                    'reinforcement': order.get('reinforcement', 0),
                                    'kp_date': order.get('kp_date', 'неизвестно'),
                                    'customer': order.get('customer', 'неизвестно'),
                                    'plate_name': order.get('plate_name', '')
                                }
                                break
                        
                        # Если не нашли в orders_2d, пробуем find_plate_info
                        if not plate_info:
                            plate_info = find_plate_info(length, width)
                        
                        # Если всё равно не нашли, используем fallback
                        if not plate_info:
                            plate_info = {
                                'reinforcement': item.get('reinforcement', 0),
                                'kp_date': 'неизвестно',
                                'customer': 'неизвестно',
                                'plate_name': ''
                            }
                        
                        # Проверяем, есть ли уже такая плита в списке
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
                    
                    # Формируем сообщение для дорожки
                    if plates_info:
                        track_message = f"📋 Дорожка {track_number}:\n\n"
                        
                        # Сортируем плиты по длине (от большей к меньшей)
                        plates_info.sort(key=lambda x: x['length'], reverse=True)
                        
                        for plate in plates_info:
                            reinforcement_str = f"{plate['reinforcement']:.1f}"
                            plate_name = plate.get('plate_name', '')
                            
                            # Формируем строку с названием плиты
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
                        
                        # Отправляем сообщение (разбиваем если слишком длинное)
                        if len(track_message) > 4000:
                            # Разбиваем на части
                            lines = track_message.split('\n')
                            current_part = lines[0] + '\n\n'  # Заголовок дорожки
                            
                            for line in lines[2:]:  # Пропускаем пустую строку после заголовка
                                if len(current_part + line + '\n') > 3900:
                                    await message.answer(current_part)
                                    current_part = line + '\n'
                                else:
                                    current_part += line + '\n'
                            
                            if current_part.strip():
                                await message.answer(current_part)
                        else:
                            await message.answer(track_message)
            
            start_index += tracks_in_file
            file_number += 1
        
        # === ШАГ 7: ФОРМИРУЕМ ФИНАЛЬНУЮ СТАТИСТИКУ ===
        # Собираем плиты только из первого файла (дорожки 1 до tracks_count)
        # Функция поиска исходных данных с приоритетом поиска из БД
        def find_source_info_for_stat(length, width, orders_list):
            """Находит исходные данные плиты с допуском по размерам (приоритет из БД)"""
            # Сначала пробуем найти по точным размерам в plate_to_kp_info
            key = (round(length, 2), width)
            if key in plate_to_kp_info:
                kp_data = plate_to_kp_info[key]
                # Ищем остальные данные в orders_2d
                for order in orders_list:
                    if (abs(order['length'] - length) < 0.1 and 
                        abs(order['width'] - width) < 50):
                        return {
                            'load_code': order['load_code'],
                            'reinforcement': order.get('reinforcement', 0),
                            'kp_date': kp_data['kp_date'],  # Из БД!
                            'customer': kp_data['customer'],  # Из БД!
                            'plate_name': kp_data.get('plate_name', order.get('plate_name', ''))  # Из БД!
                        }
            
            # Fallback: ищем в orders_list с допуском
            best_match = None
            best_diff = float('inf')
            
            for order in orders_list:
                # Допуск: длина ±0.1м, ширина ±50мм
                len_diff = abs(order['length'] - length)
                width_diff = abs(order['width'] - width)
                
                if len_diff < 0.1 and width_diff < 50:
                    total_diff = len_diff + width_diff / 1000.0
                    if total_diff < best_diff:
                        best_diff = total_diff
                        best_match = {
                            'load_code': order['load_code'],
                            'reinforcement': order.get('reinforcement', 0),
                            'kp_date': order.get('kp_date', 'неизвестно'),
                            'customer': order.get('customer', 'неизвестно'),
                            'plate_name': order.get('plate_name', '')
                        }
            
            return best_match
        
        final_plates = []
        for track_idx in range(min(tracks_count, len(all_tracks_list))):
            track = all_tracks_list[track_idx]
            for item in track['items']:
                # Извлекаем данные из item
                length = item.get('length')
                width = item.get('width', 1200)  # По умолчанию 1200мм
                if length:
                    # Ищем исходные данные для получения kp_date, reinforcement и customer
                    source_info = find_source_info_for_stat(length, width, orders_2d)
                    if source_info is None:
                        # Fallback: берём данные от ближайшей плиты по длине
                        for order in orders_2d:
                            if abs(order['length'] - length) < 0.1:
                                source_info = {
                                    'load_code': order['load_code'],
                                    'reinforcement': order['reinforcement'],
                                    'kp_date': order.get('kp_date', 'неизвестно'),
                                    'customer': order.get('customer', 'неизвестно')
                                }
                                break
                        
                        if source_info is None and orders_2d:
                            first_order = orders_2d[0]
                            source_info = {
                                'load_code': first_order['load_code'],
                                'reinforcement': first_order['reinforcement'],
                                'kp_date': first_order.get('kp_date', 'неизвестно'),
                                'customer': first_order.get('customer', 'неизвестно')
                            }
                        
                        if source_info is None:
                            source_info = {
                                'load_code': track.get('load_code', 8),
                                'reinforcement': item.get('reinforcement', 0),
                                'kp_date': 'неизвестно',
                                'customer': 'неизвестно'
                            }
                    
                    final_plates.append({
                        'length': length,
                        'width': width,
                        'qty': 1,
                        'load_code': source_info.get('load_code', track.get('load_code', 8)),
                        'reinforcement': source_info.get('reinforcement', item.get('reinforcement', 0)),
                        'kp_date': source_info.get('kp_date', 'неизвестно'),
                        'customer': source_info.get('customer', 'неизвестно')
                    })
        
        selected_plates = final_plates
        
        # === ШАГ 8: ФОРМИРУЕМ СТАТИСТИКУ ===
        # Формируем статистику по датам (только для выбранных плит)
        dates_stat_final = {}
        reinforcement_range_final = [999.0, 0.0]
        
        for plate in selected_plates:
            date_key = plate['kp_date']
            reinforcement = plate['reinforcement']
            
            if date_key not in dates_stat_final:
                dates_stat_final[date_key] = 0
            dates_stat_final[date_key] += plate['qty']
            
            if reinforcement < reinforcement_range_final[0]:
                reinforcement_range_final[0] = reinforcement
            if reinforcement > reinforcement_range_final[1]:
                reinforcement_range_final[1] = reinforcement
        
        dates_info = "\n".join([f"  • {date}: {count} шт" for date, count in sorted(dates_stat_final.items())])
        
        total_plates_count = sum(p['qty'] for p in selected_plates)
        total_length = sum(p['length'] * p['qty'] for p in selected_plates)
        
        # Считаем уникальные типы плит (по размерам length x width)
        unique_plate_types = set((round(p['length'], 2), p['width']) for p in selected_plates)
        
        # Количество дорожек в итоге
        actual_final_tracks = min(tracks_count, total_tracks_count)
        
        final_message = (
            f"✅ Планирование производства завершено!\n\n"
            f"📊 Статистика:\n"
            f"• Дорожек запланировано: {actual_final_tracks}\n"
            f"• Плит отобрано: {total_plates_count} шт\n"
            f"• Уникальных типоразмеров: {len(unique_plate_types)}\n"
            f"• Общая длина плит: {total_length:.1f}м\n\n"
            f"📅 Распределение по датам:\n"
            f"{dates_info}\n\n"
            f"🔧 Армирование: {reinforcement_range_final[0]:.1f} - {reinforcement_range_final[1]:.1f}\n\n"
        )
        
        if total_tracks_count > tracks_count:
            remaining_count = total_tracks_count - tracks_count
            final_message += (
                f"📦 Остаток на следующий день: {remaining_count} дорожек\n"
                f"💡 Они будут использованы при следующем планировании\n\n"
            )
        
        final_message += (
            f"💡 На схеме показаны только плиты на сегодня.\n"
            f"💡 Для каждой дорожки указано максимальное армирование."
        )
        
        await message.answer(final_message, reply_markup=main_menu_kb())
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при планировании производства: {str(e)}\n\n"
            "Попробуйте снова позже.",
            reply_markup=main_menu_kb()
        )
        import traceback
        traceback.print_exc()
    
    finally:
        await state.clear()

