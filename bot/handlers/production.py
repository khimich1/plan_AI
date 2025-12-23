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
                    'kp_date': kp_date.strftime('%d.%m.%Y')
                })
        
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
        # Конвертируем в формат для оптимизатора (включая kp_date для статистики)
        orders_2d = []
        for plate_data in selected_plates:
            orders_2d.append({
                'length': plate_data['length'],
                'width': plate_data['width'],
                'qty': plate_data['qty'],
                'load_code': plate_data['load_code'],
                'reinforcement': plate_data['reinforcement'],
                'kp_date': plate_data.get('kp_date', 'неизвестно')  # Для статистики
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
        
        # === ШАГ 5: FFD РАСПРЕДЕЛЕНИЕ ПО ДОРОЖКАМ ===
        from core.optimization import first_fit_decreasing, Piece
        
        # Функция поиска исходных данных с допуском по размерам
        def find_source_info(length, width, orders_list):
            """Находит исходные данные плиты с допуском по размерам"""
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
                            'reinforcement': order['reinforcement'],
                            'kp_date': order.get('kp_date', 'неизвестно')
                        }
            
            return best_match
        
        # Собираем оптимизированные плиты для FFD
        optimized_pieces = []
        piece_to_source = {}  # Маппинг piece -> исходные данные
        
        # Отладка: проверяем размеры
        orders_sizes = set((round(o['length'], 2), o['width']) for o in orders_2d)
        print(f"[DEBUG] Размеры в orders_2d: {len(orders_sizes)} уникальных")
        
        if 'plate_assignments' in optimization_result:
            assignments_sizes = set((round(a['length'], 2), a['width']) for a in optimization_result['plate_assignments'])
            print(f"[DEBUG] Размеры в plate_assignments: {len(assignments_sizes)} уникальных")
            
            matched_count = 0
            unmatched_count = 0
            
            for idx, assignment in enumerate(optimization_result['plate_assignments']):
                length = assignment['length']
                width = assignment['width']
                
                # Получаем исходные данные с допуском
                source_info = find_source_info(length, width, orders_2d)
                
                if source_info is None:
                    unmatched_count += 1
                    # Fallback: берём данные от ближайшей плиты по длине
                    # (для резанных плит ширина может сильно отличаться)
                    for order in orders_2d:
                        if abs(order['length'] - length) < 0.1:
                            source_info = {
                                'load_code': order['load_code'],
                                'reinforcement': order['reinforcement'],
                                'kp_date': order.get('kp_date', 'неизвестно')
                            }
                            break
                    
                    # Если всё ещё не нашли — используем первый попавшийся заказ
                    if source_info is None and orders_2d:
                        first_order = orders_2d[0]
                        source_info = {
                            'load_code': first_order['load_code'],
                            'reinforcement': first_order['reinforcement'],
                            'kp_date': first_order.get('kp_date', 'неизвестно')
                        }
                    
                    if source_info is None:
                        source_info = {
                            'load_code': 8,
                            'reinforcement': 999.0,
                            'kp_date': 'неизвестно'
                        }
                else:
                    matched_count += 1
                
                piece = Piece(
                    length_m=length,
                    qty=1,
                    kind='standard',
                    load_class=source_info['load_code'],
                    width_m=width / 1000.0
                )
                optimized_pieces.append(piece)
                
                # Сохраняем связь с исходными данными (включая информацию о резах)
                piece_to_source[id(piece)] = {
                    'length': length,
                    'width': width,
                    'load_code': source_info['load_code'],
                    'reinforcement': source_info['reinforcement'],
                    'kp_date': source_info['kp_date'],
                    'source': assignment.get('source', 'primary'),  # 'primary' или 'secondary'
                    'rest_width': assignment.get('rest_width'),  # Для primary резов
                    'source_rest': assignment.get('source_rest')  # Для secondary резов
                }
            
            print(f"[DEBUG] Сопоставлено: {matched_count}, не сопоставлено: {unmatched_count}")
        
        # Запускаем FFD для распределения по дорожкам
        all_ffd_tracks = first_fit_decreasing(optimized_pieces, stock_len_m=101.0)
        actual_tracks_count = len(all_ffd_tracks)
        
        print(f"[ПЛАНИРОВАНИЕ] FFD распределение: {actual_tracks_count} дорожек из {len(optimized_pieces)} плит")
        await message.answer(
            f"📊 FFD распределение:\n"
            f"• Всего получилось дорожек: {actual_tracks_count}\n"
            f"• Требуется дорожек: {tracks_count}"
        )
        
        # === ШАГ 6: БЕРЁМ ТОЛЬКО НУЖНОЕ КОЛИЧЕСТВО ДОРОЖЕК ===
        remaining_plates = []
        
        if actual_tracks_count > tracks_count:
            # Берём только первые N дорожек
            final_tracks = all_ffd_tracks[:tracks_count]
            remaining_tracks = all_ffd_tracks[tracks_count:]
            
            # Считаем статистику
            final_plates_count = sum(len(t.pieces) for t in final_tracks)
            remaining_plates_count = sum(len(t.pieces) for t in remaining_tracks)
            
            await message.answer(
                f"✂️ Фильтрация дорожек:\n"
                f"• Оставлено: {tracks_count} дорожек ({final_plates_count} плит)\n"
                f"• Отложено на потом: {actual_tracks_count - tracks_count} дорожек ({remaining_plates_count} плит)"
            )
            
            # Собираем информацию об отложенных плитах для статистики
            for track in remaining_tracks:
                for piece in track.pieces:
                    source = piece_to_source.get(id(piece), {})
                    remaining_plates.append({
                        'length': piece.length_m,
                        'width': int(piece.width_m * 1000),
                        'qty': 1,
                        'load_code': source.get('load_code', 8),
                        'reinforcement': source.get('reinforcement', 999.0),
                        'kp_date': source.get('kp_date', 'неизвестно')
                    })
        else:
            # Дорожек меньше или равно нужному — берём все
            final_tracks = all_ffd_tracks
        
        # === ШАГ 7: ФОРМИРУЕМ ДАННЫЕ ДЛЯ ВИЗУАЛИЗАЦИИ ===
        # Собираем плиты только из выбранных дорожек (final_tracks)
        final_plates = []
        for track in final_tracks:
            for piece in track.pieces:
                source = piece_to_source.get(id(piece), {})
                final_plates.append({
                    'length': piece.length_m,
                    'width': int(piece.width_m * 1000),
                    'qty': 1,
                    'load_code': source.get('load_code', 8),
                    'reinforcement': source.get('reinforcement', 999.0),
                    'kp_date': source.get('kp_date', 'неизвестно')
                })
        
        # Обновляем selected_plates для финальной статистики
        selected_plates = final_plates
        
        # === СОЗДАЁМ ОТФИЛЬТРОВАННЫЙ РЕЗУЛЬТАТ ОПТИМИЗАЦИИ ТОЛЬКО ИЗ ВЫБРАННЫХ ДОРОЖЕК ===
        # Визуализатор использует OPT_CASCADING_PLAN, поэтому нужно передать только нужные плиты
        # ВАЖНО: Сохраняем информацию о реальных резах из исходного optimization_result!
        
        # Создаём множество плит из final_tracks для быстрого поиска
        final_plates_set = set()
        for plate_data in final_plates:
            # Используем (length, width) как ключ с округлением
            key = (round(plate_data['length'], 2), plate_data['width'])
            final_plates_set.add(key)
        
        # Фильтруем plate_assignments из исходного результата, сохраняя информацию о резах
        filtered_plate_assignments = []
        for assignment in optimization_result.get('plate_assignments', []):
            length = assignment.get('length')
            width = assignment.get('width')
            if length and width:
                key = (round(length, 2), width)
                if key in final_plates_set:
                    # Находим соответствующие метаданные из piece_to_source
                    for piece_data in final_plates:
                        if (round(piece_data['length'], 2) == round(length, 2) and 
                            piece_data['width'] == width):
                            filtered_plate_assignments.append({
                                'length': length,
                                'width': width,
                                'source': assignment.get('source', 'primary'),  # Сохраняем источник!
                                'rest_width': assignment.get('rest_width'),  # Для primary
                                'source_rest': assignment.get('source_rest'),  # Для secondary
                                'load_code': piece_data['load_code'],
                                'reinforcement': piece_data['reinforcement']
                            })
                            break
        
        # Фильтруем primary_cuts из исходного результата (сохраняем реальные резы!)
        filtered_primary_cuts = []
        for cut in optimization_result.get('primary_cuts', []):
            # Проверяем, есть ли плиты из этого реза в final_plates
            filtered_lengths = []
            for length in cut.get('lengths', []):
                key = (round(length, 2), cut['width'])
                if key in final_plates_set:
                    filtered_lengths.append(length)
            
            if filtered_lengths:
                filtered_primary_cuts.append({
                    'width': cut['width'],
                    'rest': cut['rest'],
                    'qty': len(filtered_lengths),
                    'lengths': filtered_lengths
                })
        
        # Фильтруем secondary_cuts из исходного результата (сохраняем вторичные резы!)
        filtered_secondary_cuts = []
        for cut in optimization_result.get('secondary_cuts', []):
            # Получаем размеры выходных плит из вторичного реза
            output_length = cut.get('output_length')
            if not output_length and cut.get('lengths'):
                output_length = cut['lengths'][0] if isinstance(cut['lengths'], list) else None
            
            output_width = cut.get('output_width')
            if not output_width and cut.get('cuts'):
                cuts_list = cut['cuts']
                if isinstance(cuts_list, list) and len(cuts_list) > 0:
                    output_width = cuts_list[0]
            
            if output_length and output_width:
                key = (round(output_length, 2), output_width)
                # Проверяем, есть ли хотя бы одна плита из этого реза в final_plates
                if key in final_plates_set:
                    # Подсчитываем количество плит из этого реза в final_plates
                    count = sum(1 for p in final_plates 
                              if round(p['length'], 2) == round(output_length, 2) 
                              and p['width'] == output_width)
                    if count > 0:
                        # Создаём копию реза с обновлённым количеством
                        filtered_cut = cut.copy()
                        filtered_cut['qty'] = count
                        filtered_secondary_cuts.append(filtered_cut)
        
        # Создаём отфильтрованный результат оптимизации
        filtered_optimization_result = {
            'primary_cuts': filtered_primary_cuts,
            'secondary_cuts': filtered_secondary_cuts,  # ✅ Сохраняем вторичные резы!
            'total_plates': len(final_plates),
            'plate_assignments': filtered_plate_assignments,
            'total_cost': optimization_result.get('total_cost', 0)
        }
        
        # Сохраняем ОТФИЛЬТРОВАННЫЙ результат в глобальный план
        optimization.OPT_CASCADING_PLAN = filtered_optimization_result
        
        # Для совместимости с визуализацией
        all_loads = set(p['load_code'] for p in final_plates)
        filtered_optimization_result['loads_in_group'] = sorted(all_loads)
        optimization.OPT_CASCADING_PLAN_BY_LOAD = {'all': filtered_optimization_result}
        optimization.LOAD_TO_REINFORCEMENT_MAP = {
            load_code: ['all'] for load_code in all_loads
        }
        
        print(f"[ПЛАНИРОВАНИЕ] Сохранено в глобальный план: {len(final_plates)} плит для визуализации")
        
        # === ШАГ 8: ОБНОВЛЯЕМ cfg ДЛЯ ВИЗУАЛИЗАЦИИ ===
        # Очищаем старые данные
        cfg.PLATES_1_2 = []
        cfg.PLATE_LOAD_DETAILS = {}
        
        # Заполняем данными только из выбранных дорожек
        for plate_data in final_plates:
            length = plate_data['length']
            width_m = plate_data['width'] / 1000.0
            load_code = plate_data['load_code']
            
            # Добавляем в PLATE_LOAD_DETAILS
            key = (length, width_m, load_code)
            if key in cfg.PLATE_LOAD_DETAILS:
                cfg.PLATE_LOAD_DETAILS[key] += 1
            else:
                cfg.PLATE_LOAD_DETAILS[key] = 1
            
            # Добавляем в соответствующий список по ширине (для обратной совместимости)
            if abs(width_m - 1.2) < 0.01:
                cfg.PLATES_1_2.append(length)
        
        # === ШАГ 9: ГЕНЕРИРУЕМ ВИЗУАЛИЗАЦИЮ ===
        await message.answer("⏳ Генерирую схему дорожек с армированием...")
        
        result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR_STR)
        
        if isinstance(result_paths, tuple) and len(result_paths) >= 2:
            png_path, pdf_schema_path = result_paths
            
            # Отправляем PDF схемы
            if os.path.exists(pdf_schema_path):
                await message.answer_document(
                    FSInputFile(pdf_schema_path),
                    caption="📐 Схема дорожек с армированием"
                )
            
            # Ищем детальную разбивку
            base = os.path.basename(png_path)
            if 'КЗ_' in base:
                timestamp = base.split('КЗ_', 1)[-1].replace('.png', '')
            else:
                timestamp = base.rsplit('_', 1)[-1].replace('.png', '')
            
            breakdown_path = os.path.join(OUTPUTS_DIR_STR, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx')
            
            if os.path.exists(breakdown_path):
                await message.answer_document(
                    FSInputFile(breakdown_path),
                    caption="📊 Детальная разбивка компонентов"
                )
        
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
        actual_final_tracks = len(final_tracks) if 'final_tracks' in dir() else tracks_count
        
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
        
        if remaining_plates and len(remaining_plates) > 0:
            remaining_count = sum(p['qty'] for p in remaining_plates)
            final_message += (
                f"📦 Остаток на следующий день: {remaining_count} плит\n"
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

