"""Экспорт диаграммы Ганта и сохранение планов"""
import asyncio
import logging
import os
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

from core.gantt_excel import create_gantt_excel
from core import kp_db
import core.config_and_data as cfg

from ..keyboards import main_menu_kb, calendar_days_kb
from ..bot_config import OUTPUTS_DIR_STR

# Импорт менеджера планов
from .plan_manager import (
    get_active_plan_id, add_tracks_to_plan, format_plan_stats_message,
    get_all_tracks_from_plan, get_global_days_info, get_global_day_occupancy,
    MAX_TRACKS_PER_DAY, get_all_plans_gantt_data, convert_lookup_keys_to_tuples,
    save_plan, update_plan_metadata, set_active_plan, get_plan_path
)

router = Router()


# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ЗАЩИТЫ ОТ ПОТЕРИ ПЛИТ ===

def _count_plates_in_tracks(all_tracks_list: list) -> dict:
    """
    Подсчитывает плиты в tracks с группировкой по (length, width, load_code).
    
    Простыми словами:
    - Проходит по всем дорожкам и плитам
    - Считает, сколько плит каждого размера И load_code попало в план
    - ИСПРАВЛЕНИЕ: Также считает плиты из secondary_cuts (вторичных резов)
    - ИСПРАВЛЕНИЕ: Теперь ключ включает load_code для различения плит с разной нагрузкой
    - Возвращает словарь {(length, width, load_code): qty}
    """
    counts = {}  # {(length, width, load_code): qty}
    # #region agent log
    import json as _json
    _debug_log = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
    # #endregion
    
    for track in all_tracks_list:
        for item in track.get('items', []):
            if not item:
                continue
            
            length = round(item.get('length', 0), 2)
            load_code = cfg.normalize_load_code(item.get('load_code', 8))  # ИСПРАВЛЕНИЕ: получаем load_code
            
            # Определяем ширину в зависимости от mode
            mode = item.get('mode', 'solid')
            if mode == 'split':
                width = round(item.get('main_w', 1.2) * 1000)  # round для корректного округления float
            elif mode == 'transverse':
                width = round(item.get('width', 1.2) * 1000)
            else:
                # solid mode - берём width из поля (ИСПРАВЛЕНИЕ: учитываем реальную ширину)
                width = round(item.get('width', 1.2) * 1000)
            
            # ИСПРАВЛЕНИЕ: ключ теперь включает load_code
            key = (length, width, load_code)
            counts[key] = counts.get(key, 0) + 1
            
            plate_name = item.get('plate_name', '')
            # #region agent log H6: ВСЕ primary плиты с нестандартной шириной (не 1200/720)
            if width not in [1200, 720, 1080]:
                with open(_debug_log, 'a', encoding='utf-8') as _f:
                    _f.write(_json.dumps({"hypothesisId": "H6", "location": "production_export:_count_plates_in_tracks:primary", "message": "Primary плита (нестандартная)", "data": {"plate_name": plate_name, "length": length, "width": width, "load_code": load_code, "mode": mode, "key": str(key)}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
            # #endregion
            
            # DEBUG: логируем плиты с нагрузкой 16п
            if '16п' in plate_name or load_code == 1600:
                logger.debug(f"[COUNT_TRACKS] Плита: {plate_name}, длина={length}, ширина={width}, load_code={load_code}, mode={mode}")
            
            # НОВОЕ: Плиты из вторичных резов (secondary_cuts)
            # Эти плиты получены из остатков primary плит и тоже должны учитываться
            # Примечание: secondary_cuts наследуют load_code от родительской плиты
            for sec_cut in item.get('secondary_cuts', []) or []:
                sec_width = round(sec_cut.get('width', 0) * 1000)  # round для корректного округления
                # Длина: если есть target_length (поперечный рез), иначе длина родительской плиты
                sec_length = sec_cut.get('target_length') or length
                if sec_width > 0:
                    sec_key = (round(sec_length, 2), sec_width, load_code)  # наследуем load_code
                    counts[sec_key] = counts.get(sec_key, 0) + 1
                    # #region agent log H6: ВСЕ вторичные резы (не только 460/530/665)
                    with open(_debug_log, 'a', encoding='utf-8') as _f:
                        _f.write(_json.dumps({"hypothesisId": "H6", "location": "production_export:_count_plates_in_tracks:secondary", "message": "Вторичный рез", "data": {"parent_plate": plate_name, "sec_length": sec_length, "sec_width": sec_width, "load_code": load_code, "sec_key": str(sec_key), "sec_cut_raw": str(sec_cut)[:200]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                    # #endregion
    
    return counts


def _find_lost_plates(orders_2d: list, plates_in_tracks: dict, tolerance: float = 0.03) -> list:
    """
    Находит плиты из orders_2d, которые не попали в tracks.
    
    Простыми словами:
    - Для каждого заказа проверяет, сколько плит попало в tracks
    - Если попало меньше, чем заказано — плиты "потеряны"
    - Возвращает список потерянных плит для возврата в производство
    
    ИСПРАВЛЕНИЕ: Теперь учитывает load_code для различения плит с разной нагрузкой:
    - Плиты с одинаковыми размерами, но разным load_code (8п vs 16п) учитываются отдельно
    - Плита 33,8-12-16п НЕ может быть использована для выполнения заказа 33,8-12-8п
    
    Аргументы:
        orders_2d: список заказов [{'length', 'width', 'qty', 'load_code', 'kp_id', 'plate_name'}, ...]
        plates_in_tracks: словарь {(length, width, load_code): qty} — что попало в tracks
        tolerance: допуск по длине (метры) для fuzzy-поиска
    
    Возвращает:
        список потерянных плит [{'kp_id', 'plate_name', 'qty_lost'}, ...]
    """
    lost = []
    tracks_used = {}  # Отслеживаем, сколько плит уже "использовали" из tracks
    # #region agent log
    import json as _json2
    _debug_log2 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
    # #endregion
    
    for order in orders_2d:
        length = round(order.get('length', 0), 2)
        width = order.get('width', 1200)
        load_code = cfg.normalize_load_code(order.get('load_code', 8))  # ИСПРАВЛЕНИЕ: получаем load_code
        qty_ordered = order.get('qty', 1)
        kp_id = order.get('kp_id')
        plate_name = order.get('plate_name', '')
        
        # Ищем в tracks с fuzzy по длине И точным совпадением по load_code
        qty_found = 0
        
        for track_key, t_qty in plates_in_tracks.items():
            # Распаковываем ключ (может быть 2 или 3 элемента для обратной совместимости)
            if len(track_key) == 3:
                t_len, t_width, t_load_code = track_key
            else:
                t_len, t_width = track_key
                t_load_code = 8  # значение по умолчанию для старых данных
            t_load_code = cfg.normalize_load_code(t_load_code)
            
            # ИСПРАВЛЕНИЕ: Проверяем совпадение по load_code
            load_code_matches = (t_load_code == load_code)
            
            # Проверяем совпадение по ширине с tolerance 20мм
            width_matches = abs(t_width - width) <= 20
            
            if load_code_matches and width_matches and abs(t_len - length) <= tolerance:
                already_used = tracks_used.get(track_key, 0)
                available = t_qty - already_used
                
                if available > 0:
                    take = min(available, qty_ordered - qty_found)
                    qty_found += take
                    tracks_used[track_key] = already_used + take
                    
                    if qty_found >= qty_ordered:
                        break
        
        # #region agent log H2: результат поиска для плит 460, 530, 665
        if width in [460, 530, 665] or '4,6' in plate_name or '5,3' in plate_name or '6,65' in plate_name:
            with open(_debug_log2, 'a', encoding='utf-8') as _f2:
                _f2.write(_json2.dumps({"hypothesisId": "H2", "location": "production_export:_find_lost_plates", "message": "Поиск плиты в tracks", "data": {"plate_name": plate_name, "kp_id": kp_id, "length": length, "width": width, "load_code": load_code, "qty_ordered": qty_ordered, "qty_found": qty_found, "is_lost": qty_found < qty_ordered}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
        # #endregion
        
        # Если нашли меньше, чем заказано — плиты потеряны
        if qty_found < qty_ordered:
            logger.warning(f"[LOST_PLATE] Потеряна: {plate_name} x{qty_ordered - qty_found} "
                          f"(заказ: длина={length}, ширина={width}, load_code={load_code}, qty={qty_ordered}, найдено={qty_found})")
            # Логируем доступные ключи с похожей длиной
            similar_keys = [(k, v) for k, v in plates_in_tracks.items() 
                           if abs(k[0] - length) <= tolerance * 2]
            if similar_keys:
                logger.warning(f"[LOST_PLATE]   Похожие в tracks: {similar_keys[:5]}")
            lost.append({
                'kp_id': kp_id,
                'plate_name': plate_name,
                'qty_lost': qty_ordered - qty_found,
                'load_code': load_code  # ИСПРАВЛЕНИЕ: сохраняем load_code для отладки
            })
    
    return lost


@router.callback_query(F.data == "export_gantt")
async def export_gantt_chart(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "Диаграмма Ганта".
    Создаёт СУММАРНУЮ Excel-диаграмму по ВСЕМ сохранённым планам производства.
    
    Простыми словами:
    - Загружает ВСЕ сохранённые планы
    - Собирает все дорожки и информацию о КП
    - Создаёт одну большую диаграмму Ганта по всем планам
    """
    await callback.message.answer("📊 Создаю суммарную диаграмму Ганта по всем планам...")
    
    # Получаем данные из ВСЕХ планов
    gantt_data = get_all_plans_gantt_data()
    
    if not gantt_data:
        await callback.message.answer(
            "❌ Нет сохранённых планов для создания диаграммы.\n\n"
            "💡 Сначала создайте и сохраните план:\n"
            "1️⃣ Нажмите «🚀 Начать планирование»\n"
            "2️⃣ Выберите КП для производства\n"
            "3️⃣ Нажмите «💾 Сохранить план»",
            reply_markup=main_menu_kb()
        )
        await callback.answer()
        return
    
    # Извлекаем данные
    all_tracks_list = gantt_data['all_tracks']
    plate_lookup_exact = gantt_data['plate_lookup_exact']
    plate_lookup_by_length = gantt_data['plate_lookup_by_length']
    start_date_for_gantt = gantt_data['earliest_start_date']
    plans_count = gantt_data['plans_count']
    total_days = gantt_data['total_days']
    
    # Для корректного подсчёта дней используем среднее количество дорожек
    # (это нужно для совместимости с create_gantt_excel)
    tracks_count = 3  # Среднее значение, не критично для диаграммы
    
    try:
        # Создаём диаграмму Ганта
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_list,
            tracks_count=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            output_dir=OUTPUTS_DIR_STR,
            start_date=start_date_for_gantt
        )
        
        if gantt_path and os.path.exists(gantt_path):
            # Форматируем даты для отображения
            start_date_str = start_date_for_gantt.strftime('%d.%m.%Y')
            end_date_str = gantt_data['latest_end_date'].strftime('%d.%m.%Y')
            
            # Считаем количество уникальных КП в диаграмме
            # (это можно сделать только после создания файла, но для простоты опустим)
            
            await callback.message.answer_document(
                FSInputFile(gantt_path),
                caption=(
                    "📊 СУММАРНАЯ диаграмма Ганта по всем планам\n\n"
                    f"📅 Период: {start_date_str} — {end_date_str}\n"
                    f"📋 Планов: {plans_count}\n"
                    f"📆 Дней: {total_days}\n"
                    f"🛤️ Дорожек: {len(all_tracks_list)}\n\n"
                    "Цветовая кодировка:\n"
                    "🟢 Зелёный — успеваем до дедлайна\n"
                    "🟡 Жёлтый — завершаем в день дедлайна\n"
                    "🔴 Красный — опаздываем!"
                )
            )
            
            logger.info(f"[GANTT] Диаграмма успешно создана: {gantt_path}")
        else:
            await callback.message.answer(
                "⚠️ Не удалось создать диаграмму.\n"
                "Возможно, нет данных о КП в сохранённых планах.\n\n"
                "💡 Убедитесь, что планы содержат информацию о заказах."
            )
    
    except Exception as e:
        logger.exception(f"Ошибка создания диаграммы: {e}")
        await callback.message.answer(
            "❌ Не удалось создать диаграмму Ганта.\n"
            "Подробности в logs/bot.log."
        )
    
    # Получаем данные из state для возврата к календарю
    data = await state.get_data()
    total_days_state = data.get('total_days', 1)
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = data.get('days_info', {})
    from_saved_plan = data.get('from_saved_plan', False)
    
    # Показываем клавиатуру выбора дней снова (с датами)
    # только если мы в контексте просмотра календаря
    if total_days_state and total_days_state > 0:
        await callback.message.answer(
            "Выберите день для просмотра:",
            reply_markup=calendar_days_kb(
                total_days_state, 
                plan_start_date, 
                completed_days,
                days_info,
                show_save_button=not from_saved_plan
            )
        )
    
    await callback.answer()


@router.callback_query(F.data == "save_current_plan")
async def save_current_plan(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик сохранения актуального плана производства.
    
    НОВАЯ ЛОГИКА:
    - Добавляет дорожки к активному плану (не перезаписывает!)
    - Если активного плана нет — создаёт новый
    - Показывает статистику: какие дни обновлены, какие созданы
    - Генерирует диаграмму Ганта
    - БЛОКИРУЕТ сохранение при превышении лимита дорожек
    """
    # Получаем данные из state
    data = await state.get_data()
    all_tracks_list = data.get('all_tracks_list', [])
    tracks_count = data.get('tracks_count', 1)
    plate_lookup_exact = data.get('plate_lookup_exact', {})
    plate_lookup_by_length = data.get('plate_lookup_by_length', {})
    
    # Дата начала плана
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    
    # Дополнительные данные для плана
    orders_2d = data.get('orders_2d', [])
    optimization_result = data.get('optimization_result', {})
    
    if not all_tracks_list:
        await callback.message.answer(
            "❌ Нет данных для сохранения.\n"
            "Сначала выполните анализ производства.",
            reply_markup=main_menu_kb()
        )
        await callback.answer()
        return
    
    # === ПРОВЕРКА ПРЕВЫШЕНИЯ ЛИМИТА ===
    # ИСПРАВЛЕНИЕ: Берём ID плана ТОЛЬКО из state, без fallback на глобальный!
    # Это гарантирует создание нового плана при active_plan_id=None
    active_plan_id = data.get('active_plan_id')  # Может быть None — это нормально!
    
    # Исключаем текущий план из подсчёта занятости (чтобы не считать дважды)
    # Это исправляет баг, когда при просмотре существующего плана и попытке сохранения
    # система считала дорожки текущего плана дважды
    global_occupancy = get_global_day_occupancy(exclude_plan_id=active_plan_id)
    
    total_days = data.get('total_days', 1)
    try:
        start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
    except:
        start_dt = datetime.now()
    
    overloaded_days = []
    for day_num in range(1, total_days + 1):
        day_date = start_dt + timedelta(days=day_num - 1)
        date_key = day_date.strftime('%Y-%m-%d')
        date_display = day_date.strftime('%d.%m')
        
        current_occupied = global_occupancy.get(date_key, 0)
        free_slots = MAX_TRACKS_PER_DAY - current_occupied
        
        if tracks_count > free_slots:
            overloaded_days.append({
                'date': date_display,
                'occupied': current_occupied,
                'free': free_slots,
                'want': tracks_count
            })
    
    # Если есть превышение - БЛОКИРУЕМ сохранение
    if overloaded_days:
        error_lines = ["❌ НЕЛЬЗЯ СОХРАНИТЬ ПЛАН!\n"]
        error_lines.append(f"Превышен лимит дорожек ({MAX_TRACKS_PER_DAY}/день):\n")
        
        for day in overloaded_days[:5]:
            error_lines.append(
                f"  • {day['date']}: занято {day['occupied']}/5, "
                f"свободно {day['free']}, нужно {day['want']}"
            )
        
        if len(overloaded_days) > 5:
            error_lines.append(f"  ... и ещё {len(overloaded_days) - 5} дней")
        
        error_lines.append(f"\n💡 Что делать:")
        error_lines.append(f"1️⃣ Начните планирование заново с меньшим кол-вом дорожек")
        error_lines.append(f"2️⃣ Или выберите другую дату начала")
        error_lines.append(f"3️⃣ Или удалите/отредактируйте другие планы")
        
        await callback.message.answer('\n'.join(error_lines))
        await callback.answer("⚠️ Превышен лимит!")
        return
    
    await callback.message.answer("💾 Сохраняю дорожки в план...")
    
    plan_saved = False  # Флаг для отката
    plan_id = None
    
    try:
        
        # ШАБЛОН ИСПРАВЛЕНИЯ: Подготавливаем план БЕЗ сохранения на диск
        updated_plan, stats = add_tracks_to_plan(
            plan_id=active_plan_id,
            new_tracks_list=all_tracks_list,
            start_date=plan_start_date,
            tracks_per_day=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            orders_2d=orders_2d,
            optimization_result=optimization_result,
            auto_save=False  # НЕ сохраняем автоматически!
        )
        
        db_path = str(PROJECT_ROOT / "plita.db")
        plan_id = updated_plan['id']
        
        # === ИСПРАВЛЕНО: СНАЧАЛА ОПРЕДЕЛЯЕМ ПОТЕРЯННЫЕ ПЛИТЫ ===
        # Подсчитываем, сколько плит реально попало в tracks ПЕРЕД пометкой
        plates_in_tracks = _count_plates_in_tracks(all_tracks_list)
        lost_plates = _find_lost_plates(orders_2d, plates_in_tracks, tolerance=0.03)
        
        # Создаём множество потерянных плит для быстрой проверки
        lost_plates_set = set()
        if lost_plates:
            for lp in lost_plates:
                # Ключ: (kp_id, plate_name) для идентификации потерянной плиты
                lost_plates_set.add((lp.get('kp_id'), lp.get('plate_name')))
            lost_info = ", ".join([f"{lp['plate_name']} x{lp['qty_lost']}" for lp in lost_plates[:3]])
            if len(lost_plates) > 3:
                lost_info += f" и ещё {len(lost_plates) - 3}..."
            logger.warning(f"[SAVE_PLAN] Обнаружены потерянные плиты (НЕ будут помечены как 'в плане'): {lost_info}")
        
        # === ТЕПЕРЬ ПОМЕЧАЕМ ТОЛЬКО ПЛИТЫ, КОТОРЫЕ В ПЛАНЕ ===
        # НЕ помечаем потерянные плиты - они остаются "в производстве"
        plates_marked = 0
        plates_failed = 0
        plates_skipped = 0  # Пропущенные (потерянные)
        
        for order in orders_2d:
            kp_id = order.get('kp_id')
            plate_name = order.get('plate_name')
            qty = order.get('qty', 1)
            
            # Проверяем, не потеряна ли эта плита
            order_key = (kp_id, plate_name)
            
            if order_key in lost_plates_set:
                # Находим, сколько плит потеряно для этого заказа
                qty_lost = 0
                for lp in lost_plates:
                    if lp.get('kp_id') == kp_id and lp.get('plate_name') == plate_name:
                        qty_lost = lp.get('qty_lost', qty)
                        break
                
                qty_to_mark = qty - qty_lost
                if qty_to_mark <= 0:
                    plates_skipped += 1
                    logger.info(f"[SAVE_PLAN] Пропускаем потерянную плиту: КП #{kp_id}, {plate_name} x{qty} (вся потеряна)")
                    continue
                else:
                    # Частичная потеря - помечаем только то, что попало в план
                    logger.info(f"[SAVE_PLAN] Частичная потеря: КП #{kp_id}, {plate_name} - помечаем {qty_to_mark} из {qty}")
                    qty = qty_to_mark
            
            if kp_id and plate_name and qty > 0:
                success = kp_db.mark_plates_as_planned(
                    kp_id=kp_id,
                    plate_name=plate_name,
                    qty_to_plan=qty,
                    plan_id=plan_id,
                    db_path=db_path
                )
                if success:
                    plates_marked += 1
                else:
                    plates_failed += 1
                    logger.warning(f"[SAVE_PLAN] Не удалось пометить плиту: КП #{kp_id}, {plate_name} x{qty}")
        
        logger.info(f"[SAVE_PLAN] Помечено {plates_marked} позиций плит как 'в плане' для плана {plan_id}")
        if plates_skipped > 0:
            logger.info(f"[SAVE_PLAN] Пропущено {plates_skipped} потерянных плит (остаются 'в производстве')")
        
        # Если не удалось пометить плиты - откатываем и не сохраняем план
        if plates_failed > 0:
            logger.error(f"[SAVE_PLAN] Не удалось пометить {plates_failed} плит. Откатываю помеченные плиты...")
            kp_db.return_plan_plates_to_production(plan_id, db_path)
            raise Exception(f"Не удалось пометить {plates_failed} плит в БД")
        
        # === ТЕПЕРЬ СОХРАНЯЕМ ПЛАН НА ДИСК ===
        # Плиты уже помечены в БД, теперь можно безопасно сохранить план
        save_plan(updated_plan)
        update_plan_metadata(updated_plan)
        set_active_plan(plan_id)
        plan_saved = True  # Отмечаем, что план сохранён
        logger.info(f"[SAVE_PLAN] План {plan_id} успешно сохранён на диск")
        
        # Сохраняем ID активного плана в state и устанавливаем флаг from_saved_plan
        await state.update_data(
            active_plan_id=updated_plan['id'],
            from_saved_plan=True  # План теперь сохранён
        )
        
        # Парсим дату начала для диаграммы Ганта
        start_date_for_gantt = datetime.now()
        if plan_start_date:
            try:
                start_date_for_gantt = datetime.strptime(plan_start_date, '%Y-%m-%d')
            except ValueError:
                pass
        
        # Собираем все дорожки из обновлённого плана для диаграммы Ганта
        all_tracks_for_gantt = get_all_tracks_from_plan(updated_plan)
        
        # Создаём диаграмму Ганта
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_for_gantt,
            tracks_count=tracks_count,
            plate_lookup_exact=convert_lookup_keys_to_tuples(updated_plan.get('plate_lookup_exact', {})),
            plate_lookup_by_length=convert_lookup_keys_to_tuples(updated_plan.get('plate_lookup_by_length', {})),
            output_dir=OUTPUTS_DIR_STR,
            start_date=start_date_for_gantt
        )
        
        # Формируем сообщение со статистикой
        stats_message = format_plan_stats_message(stats)
        
        # Подсчитываем общую статистику плана
        total_days = len(updated_plan.get('days', {}))
        total_tracks = sum(
            day.get('saved_tracks_count', len(day.get('tracks', [])))
            for day in updated_plan.get('days', {}).values()
        )
        
        # Форматируем дату начала для отображения
        start_date_display = plan_start_date
        try:
            start_date_display = datetime.strptime(plan_start_date, '%Y-%m-%d').strftime('%d.%m.%Y')
        except:
            pass
        
        # Сообщение об успешном сохранении
        success_message = (
            f"✅ План успешно сохранён!\n\n"
            f"{stats_message}\n\n"
            f"📋 План: {updated_plan.get('name', 'Без названия')}\n"
            f"📅 Дата начала: {start_date_display}\n"
            f"📊 Всего дней: {total_days}\n"
            f"🛤️ Всего дорожек: {total_tracks}\n\n"
            f"⭐ Этот план установлен как АКТИВНЫЙ\n"
            f"При входе в «Календарный план» из меню откроется именно он.\n\n"
            f"💡 Как открыть:\n"
            f"Планирование производства → Календарный план"
        )
        
        await callback.message.answer(success_message)
        
        # Получаем информацию о днях с ГЛОБАЛЬНОЙ загруженностью
        days_info = get_global_days_info(updated_plan)
        completed_days = updated_plan.get('completed_days', [])
        
        # Показываем обновлённый календарь (план уже сохранён, кнопка не нужна)
        await callback.message.answer(
            "📅 Обновлённый календарь:",
            reply_markup=calendar_days_kb(
                total_days, 
                plan_start_date, 
                completed_days, 
                days_info,
                show_save_button=False  # План только что сохранён
            )
        )
        
    except Exception as e:
        logger.exception(f"Ошибка при сохранении плана: {e}")
        
        # === ОТКАТ ИЗМЕНЕНИЙ ===
        # Если плиты были помечены, но сохранение плана не удалось - возвращаем плиты
        if plan_id:
            try:
                db_path = str(PROJECT_ROOT / "plita.db")
                recovered = kp_db.return_plan_plates_to_production(plan_id, db_path)
                if recovered > 0:
                    logger.info(f"[ROLLBACK] Возвращено {recovered} плит в производство")
            except Exception as rollback_error:
                logger.error(f"[ROLLBACK] Ошибка при откате плит: {rollback_error}")
        
        # Если план был сохранён на диск, но потом произошла ошибка - удаляем файл
        if plan_saved and plan_id:
            try:
                plan_path = get_plan_path(plan_id)
                if plan_path.exists():
                    os.remove(plan_path)
                    logger.info(f"[ROLLBACK] Удалён файл плана {plan_id}")
            except Exception as delete_error:
                logger.error(f"[ROLLBACK] Ошибка при удалении файла плана: {delete_error}")
        
        await callback.message.answer(
            "❌ Не удалось сохранить план.\n"
            "Все изменения отменены.\n"
            "Подробности в logs/bot.log."
        )
    
    await callback.answer()
