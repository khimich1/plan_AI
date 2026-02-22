"""Экспорт диаграммы Ганта и сохранение планов"""
import asyncio
import logging
import os
from datetime import datetime, timedelta
from typing import Tuple
from pathlib import Path

logger = logging.getLogger(__name__)
_DEBUG_SESSION_LOG = r"c:\Users\Роман\Desktop\Шишов\debug-d7e22e.log"

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

from ..keyboards import main_menu_kb, calendar_days_kb, production_menu_kb
from ..bot_config import OUTPUTS_DIR_STR

# Импорт менеджера планов
from .plan_manager import (
    get_active_plan_id, add_tracks_to_plan, format_plan_stats_message,
    get_all_tracks_from_plan, get_global_days_info, get_global_day_occupancy,
    MAX_TRACKS_PER_DAY, get_all_plans_gantt_data, convert_lookup_keys_to_tuples,
    save_plan, update_plan_metadata, set_active_plan, get_plan_path
)

router = Router()


def _debug_session_write(run_id, hypothesis_id, location, message, data):
    """Пишет NDJSON в debug-d7e22e.log для Debug Mode."""
    try:
        line = __import__("json").dumps({
            "sessionId": "d7e22e",
            "runId": run_id,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(__import__("time").time() * 1000),
        }, ensure_ascii=False) + "\n"
        with open(_DEBUG_SESSION_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


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
            # Эти плиты получены из остатков primary плит; load_code берём из sec_cut (целевой заказ), иначе из родителя
            for sec_cut in item.get('secondary_cuts', []) or []:
                sec_width = round(sec_cut.get('width', 0) * 1000)  # round для корректного округления
                # Длина: если есть target_length (поперечный рез), иначе длина родительской плиты
                sec_length = sec_cut.get('target_length') or length
                if sec_width > 0:
                    sec_load_code = cfg.normalize_load_code(sec_cut.get('load_code', item.get('load_code', 8)))
                    sec_key = (round(sec_length, 2), sec_width, sec_load_code)
                    counts[sec_key] = counts.get(sec_key, 0) + 1
                    # #region agent log H6: ВСЕ вторичные резы (не только 460/530/665)
                    with open(_debug_log, 'a', encoding='utf-8') as _f:
                        _f.write(_json.dumps({"hypothesisId": "H6", "location": "production_export:_count_plates_in_tracks:secondary", "message": "Вторичный рез", "data": {"parent_plate": plate_name, "sec_length": sec_length, "sec_width": sec_width, "load_code": sec_load_code, "sec_key": str(sec_key), "sec_cut_raw": str(sec_cut)[:200]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                    # #endregion
    
    # #region agent log H_tracks: ключи в треках для длин 5.98, 5.99, 6.11, 6.12 и load 8
    _target_lengths = (5.98, 5.99, 6.11, 6.12)
    _keys_61_59 = [(k, counts.get(k, 0)) for k in counts if len(k) >= 2 and round(k[0], 2) in _target_lengths and (len(k) == 2 or cfg.normalize_load_code(k[2] if len(k) > 2 else 8) == 8)]
    if _keys_61_59:
        try:
            with open(_debug_log, 'a', encoding='utf-8') as _f:
                _f.write(_json.dumps({"hypothesisId": "H_tracks", "location": "production_export:_count_plates_in_tracks", "message": "Ключи в треках для 5.98/5.99/6.11/6.12 load 8", "data": {"keys_with_qty": [list(k) + [v] for k, v in _keys_61_59]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
        except Exception:
            pass
    # #endregion
    return counts


def _get_qty_found_for_order(order: dict, plates_in_tracks: dict, tracks_used: dict, tolerance: float) -> int:
    """
    Подсчитывает, сколько плит данного заказа реально попало в tracks.
    Обновляет tracks_used «на месте».
    """
    length = round(order.get('length', 0), 2)
    width = order.get('width', 1200)
    load_code = cfg.normalize_load_code(order.get('load_code', 8))
    qty_ordered = order.get('qty', 1)
    qty_found = 0

    for track_key, t_qty in plates_in_tracks.items():
        if len(track_key) == 3:
            t_len, t_width, t_load_code = track_key
        else:
            t_len, t_width = track_key
            t_load_code = 8
        t_load_code = cfg.normalize_load_code(t_load_code)
        if t_load_code != load_code:
            continue
        if abs(t_width - width) > 20:
            continue
        if abs(t_len - length) > tolerance:
            continue
        already_used = tracks_used.get(track_key, 0)
        available = t_qty - already_used
        if available <= 0:
            continue
        take = min(available, qty_ordered - qty_found)
        qty_found += take
        tracks_used[track_key] = already_used + take
        if qty_found >= qty_ordered:
            break
    return qty_found


def _find_lost_plates(orders_2d: list, plates_in_tracks: dict, tolerance: float = 0.03) -> Tuple[list, list]:
    """
    Находит плиты из orders_2d, которые не попали в tracks.
    
    Возвращает:
        - lost: список потерянных плит [{'kp_id', 'plate_name', 'qty_lost'}, ...]
        - orders_with_qty: список (order, qty_to_mark) для каждого заказа
                         qty_to_mark = сколько плит реально в tracks (для пометки «в плане»)
    
    ИСПРАВЛЕНИЕ: Теперь возвращает qty_to_mark для каждого заказа отдельно.
    При нескольких заказах с одним (kp_id, plate_name) qty_lost применяется к правильному заказу.
    """
    lost = []
    orders_with_qty = []  # [(order, qty_to_mark), ...]
    tracks_used = {}

    import json as _json2
    _debug_log2 = r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log"
    _agent_log = PROJECT_ROOT / "debug-73b708.log"
    # #region agent log (session 73b708) H_B: вход в _find_lost_plates
    try:
        with open(_agent_log, 'a', encoding='utf-8') as _fa:
            _fa.write(_json2.dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_B", "location": "production_export:_find_lost_plates:entry", "message": "find_lost_plates entry", "data": {"orders_count": len(orders_2d), "tracks_keys_count": len(plates_in_tracks), "tracks_total_plates": sum(plates_in_tracks.values()) if plates_in_tracks else 0}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion

    for order in orders_2d:
        qty_ordered = order.get('qty', 1)
        qty_found = _get_qty_found_for_order(order, plates_in_tracks, tracks_used, tolerance)
        qty_to_mark = qty_found
        orders_with_qty.append((order, qty_to_mark))

        length = round(order.get('length', 0), 2)
        width = order.get('width', 1200)
        load_code = cfg.normalize_load_code(order.get('load_code', 8))
        kp_id = order.get('kp_id')
        plate_name = order.get('plate_name', '')

        if width in [460, 530, 665] or '4,6' in plate_name or '5,3' in plate_name or '6,65' in plate_name:
            try:
                with open(_debug_log2, 'a', encoding='utf-8') as _f2:
                    _f2.write(_json2.dumps({"hypothesisId": "H2", "location": "production_export:_find_lost_plates", "message": "Поиск плиты в tracks", "data": {"plate_name": plate_name, "kp_id": kp_id, "length": length, "width": width, "load_code": load_code, "qty_ordered": qty_ordered, "qty_found": qty_found, "is_lost": qty_found < qty_ordered}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
            except Exception:
                pass
        # #region agent log H_save: для 61,2 и 59,8 — qty в треках vs заказ
        if '61,2' in (plate_name or '') or '59,8' in (plate_name or ''):
            try:
                with open(_debug_log2, 'a', encoding='utf-8') as _f2:
                    _f2.write(_json2.dumps({"hypothesisId": "H_save", "location": "production_export:_find_lost_plates", "message": "61,2/59,8: найдено в треках vs заказ", "data": {"plate_name": plate_name, "kp_id": kp_id, "length": length, "width": width, "qty_ordered": qty_ordered, "qty_found": qty_found, "qty_to_mark": qty_to_mark, "is_lost": qty_found < qty_ordered}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
            except Exception:
                pass
        # #endregion

        if qty_found < qty_ordered:
            logger.warning(f"[LOST_PLATE] Потеряна: {plate_name} x{qty_ordered - qty_found} "
                          f"(заказ: длина={length}, ширина={width}, load_code={load_code}, qty={qty_ordered}, найдено={qty_found})")
            similar_keys = [(k, v) for k, v in plates_in_tracks.items() if abs(k[0] - length) <= tolerance * 2]
            if similar_keys:
                logger.warning(f"[LOST_PLATE]   Похожие в tracks: {similar_keys[:5]}")
            lost.append({
                'kp_id': kp_id,
                'plate_name': plate_name,
                'qty_lost': qty_ordered - qty_found,
                'load_code': load_code
            })
            # #region agent log (session 73b708) H_B: одна потерянная плита
            try:
                with open(_agent_log, 'a', encoding='utf-8') as _fa:
                    _fa.write(_json2.dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_B", "location": "production_export:_find_lost_plates:lost", "message": "Lost plate", "data": {"plate_name": plate_name, "kp_id": kp_id, "length": length, "width": width, "load_code": load_code, "qty_ordered": qty_ordered, "qty_found": qty_found, "qty_lost": qty_ordered - qty_found}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion

    # #region agent log (session 73b708) H_B: выход из _find_lost_plates
    try:
        with open(_agent_log, 'a', encoding='utf-8') as _fa:
            _fa.write(_json2.dumps({"sessionId": "73b708", "runId": "run1", "hypothesisId": "H_B", "location": "production_export:_find_lost_plates:exit", "message": "find_lost_plates exit", "data": {"lost_count": len(lost), "lost_list": [{"plate_name": x.get("plate_name"), "qty_lost": x.get("qty_lost")} for x in lost]}, "timestamp": __import__('time').time()}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    return lost, orders_with_qty


@router.callback_query(F.data == "export_gantt_current")
async def export_gantt_current_plan(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "📈 Диаграмма этого плана".
    Строит диаграмму Ганта по ТЕКУЩЕМУ плану в памяти (FSM state), даже если план ещё не сохранён.
    """
    await callback.message.answer("📈 Создаю диаграмму Ганта по текущему плану...")

    data = await state.get_data()

    all_tracks_list = data.get('all_tracks_list') or []
    tracks_count = data.get('tracks_count')
    plate_lookup_exact = data.get('plate_lookup_exact') or {}
    plate_lookup_by_length = data.get('plate_lookup_by_length') or {}
    plan_start_date = data.get('plan_start_date')

    if not all_tracks_list or not tracks_count:
        await callback.message.answer(
            "❌ Текущий план не найден в памяти.\n\n"
            "💡 Сначала выполните планирование:\n"
            "1️⃣ Планирование производства → «🚀 Начать планирование»\n"
            "2️⃣ Выберите КП и дождитесь сообщения «✅ План готов!»\n"
            "3️⃣ Затем нажмите «📈 Диаграмма этого плана»"
        )
        await callback.answer()
        return

    # Парсим дату начала (нужна для дат в Excel)
    start_date_for_gantt = datetime.now()
    if plan_start_date:
        try:
            start_date_for_gantt = datetime.strptime(str(plan_start_date)[:10], '%Y-%m-%d')
        except Exception:
            pass

    try:
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_list,
            tracks_count=int(tracks_count),
            plate_lookup_exact=convert_lookup_keys_to_tuples(plate_lookup_exact),
            plate_lookup_by_length=convert_lookup_keys_to_tuples(plate_lookup_by_length),
            output_dir=OUTPUTS_DIR_STR,
            start_date=start_date_for_gantt
        )

        if gantt_path and os.path.exists(gantt_path):
            total_days = data.get('total_days')
            total_tracks_count = data.get('total_tracks_count') or len(all_tracks_list)

            start_str = start_date_for_gantt.strftime('%d.%m.%Y')
            end_str = ''
            if isinstance(total_days, int) and total_days > 0:
                end_dt = start_date_for_gantt + timedelta(days=total_days - 1)
                end_str = end_dt.strftime('%d.%m.%Y')

            caption_lines = [
                "📈 Диаграмма Ганта этого плана (ещё не сохранённого)\n",
                f"📅 Дата начала: {start_str}",
            ]
            if end_str:
                caption_lines.append(f"📅 Период: {start_str} — {end_str}")
            if total_days:
                caption_lines.append(f"📆 Дней: {total_days}")
            caption_lines.append(f"🛤️ Дорожек: {total_tracks_count}")
            caption_lines.append("\nПодсказка:")
            caption_lines.append("• «📊 Диаграмма Ганта» — это суммарно по ВСЕМ сохранённым планам")
            caption = "\n".join(caption_lines)

            await callback.message.answer_document(
                FSInputFile(gantt_path),
                caption=caption
            )
        else:
            await callback.message.answer(
                "⚠️ Не удалось создать диаграмму.\n"
                "Возможно, в текущем плане нет данных для построения."
            )

    except Exception as e:
        logger.exception(f"Ошибка создания диаграммы текущего плана: {e}")
        await callback.message.answer(
            "❌ Не удалось создать диаграмму Ганта по текущему плану.\n"
            "Подробности в logs/bot.log."
        )

    # Возвращаем клавиатуру выбора дней (чтобы не теряться)
    total_days_state = data.get('total_days', 0)
    plan_start_date_state = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = data.get('days_info', {})
    from_saved_plan = data.get('from_saved_plan', False)

    if isinstance(total_days_state, int) and total_days_state > 0:
        await callback.message.answer(
            "Выберите день для просмотра:",
            reply_markup=calendar_days_kb(
                total_days_state,
                plan_start_date_state,
                completed_days,
                days_info,
                show_save_button=not from_saved_plan
            )
        )

    await callback.answer()


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
        
        # === ИСПРАВЛЕНО: СНАЧАЛА ОПРЕДЕЛЯЕМ ПОТЕРЯННЫЕ ПЛИТЫ И QTY_TO_MARK ===
        # Подсчитываем, сколько плит реально попало в tracks ПЕРЕД пометкой
        # ИСПРАВЛЕНИЕ: Используем qty_to_mark (qty_found) для КАЖДОГО заказа отдельно.
        # При нескольких заказах с одним (kp_id, plate_name) старый код ошибочно
        # вычитал qty_lost из всех — теперь помечаем только то, что реально в треках.
        plates_in_tracks = _count_plates_in_tracks(all_tracks_list)
        # #region agent log H_366_save: есть ли ключ (3.66, 665, 8) в треках и что в orders_2d для КП 2
        import json as _j366
        _k366 = (3.66, 665, 8)
        _in_tracks_366 = plates_in_tracks.get(_k366, 0)
        _orders_kp2 = [{"plate_name": o.get("plate_name"), "kp_id": o.get("kp_id"), "length": o.get("length"), "width": o.get("width"), "qty": o.get("qty", 1)} for o in orders_2d if o.get("kp_id") == 2]
        try:
            with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", "a", encoding="utf-8") as _f366:
                _f366.write(_j366.dumps({"hypothesisId": "H_366_save", "location": "production_export:save_plan", "message": "Перед _find_lost_plates: треки и orders_2d для КП 2", "data": {"key_366_665_8_in_tracks": _in_tracks_366, "orders_2d_kp2": _orders_kp2, "orders_2d_len": len(orders_2d)}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        # #region agent log H_viz: плиты 61,2/61,1/59,8/59,9 в треках (что пойдёт в визуализацию)
        import json as _jviz
        _viz_61_59 = []
        for tr in all_tracks_list:
            for it in tr.get('items', []) or []:
                n = (it.get('plate_name') or '')
                if any(x in n for x in ('61,2', '61,1', '59,8', '59,9')):
                    _viz_61_59.append({"plate_name": n, "length": it.get("length"), "width": it.get("width"), "kp_id": it.get("kp_id")})
        try:
            with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", "a", encoding="utf-8") as _fv:
                _fv.write(_jviz.dumps({"hypothesisId": "H_viz", "location": "production_export:save_plan", "message": "Плиты 61,2/61,1/59,8/59,9 в треках (визуализация)", "data": {"count": len(_viz_61_59), "entries": _viz_61_59[:50]}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        lost_plates, orders_with_qty = _find_lost_plates(orders_2d, plates_in_tracks, tolerance=0.03)
        # #region agent log
        _targets = []
        for _order, _qty_to_mark in orders_with_qty:
            _name = (_order.get("plate_name") or "")
            if any(_k in _name for _k in ("59,8-12-8п", "50,8-5,3-8п", "50,8-3,2-8п", "63,9-12-8п")):
                _targets.append({
                    "plate_name": _name,
                    "kp_id": _order.get("kp_id"),
                    "length": _order.get("length"),
                    "width": _order.get("width"),
                    "qty_ordered": _order.get("qty", 1),
                    "qty_to_mark": _qty_to_mark,
                })
        _debug_session_write(
            "run1",
            "H4",
            "production_export:save_plan_after_find_lost",
            "Target plates qty_to_mark after tracks matching",
            {
                "targets": _targets,
                "targets_count": len(_targets),
                "lost_targets": [
                    x for x in lost_plates
                    if any(_k in (x.get("plate_name") or "") for _k in ("59,8-12-8п", "50,8-5,3-8п", "50,8-3,2-8п", "63,9-12-8п"))
                ],
            },
        )
        # #endregion
        # #region agent log H_366_qty: qty_to_mark для плит КП 2 (36,6-6,65)
        _with_qty_kp2 = [(o.get("plate_name"), q) for o, q in orders_with_qty if o.get("kp_id") == 2]
        try:
            with open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", "a", encoding="utf-8") as _fq:
                _fq.write(_j366.dumps({"hypothesisId": "H_366_qty", "location": "production_export:save_plan", "message": "orders_with_qty для КП 2", "data": {"orders_with_qty_kp2": _with_qty_kp2}, "timestamp": __import__("time").time()}, ensure_ascii=False) + "\n")
        except Exception:
            pass
        # #endregion
        if lost_plates:
            lost_info = ", ".join([f"{lp['plate_name']} x{lp['qty_lost']}" for lp in lost_plates[:3]])
            if len(lost_plates) > 3:
                lost_info += f" и ещё {len(lost_plates) - 3}..."
            logger.warning(f"[SAVE_PLAN] Обнаружены потерянные плиты (НЕ будут помечены как 'в плане'): {lost_info}")
        
        # === ПОМЕЧАЕМ ТОЛЬКО ПЛИТЫ, КОТОРЫЕ РЕАЛЬНО В ТРЕКАХ ===
        plates_marked = 0
        plates_failed = 0
        plates_skipped = 0
        
        for order, qty_to_mark in orders_with_qty:
            kp_id = order.get('kp_id')
            plate_name = order.get('plate_name')
            qty_ordered = order.get('qty', 1)
            
            if qty_to_mark <= 0:
                if qty_ordered > 0:
                    plates_skipped += 1
                    logger.info(f"[SAVE_PLAN] Пропускаем потерянную плиту: КП #{kp_id}, {plate_name} x{qty_ordered} (вся потеряна)")
                continue
            
            if qty_to_mark < qty_ordered:
                logger.info(f"[SAVE_PLAN] Частичная потеря: КП #{kp_id}, {plate_name} - помечаем {qty_to_mark} из {qty_ordered}")
            
            if kp_id and plate_name and qty_to_mark > 0:
                success = kp_db.mark_plates_as_planned(
                    kp_id=kp_id,
                    plate_name=plate_name,
                    qty_to_plan=qty_to_mark,
                    plan_id=plan_id,
                    db_path=db_path
                )
                if success:
                    plates_marked += 1
                else:
                    plates_failed += 1
                    logger.warning(f"[SAVE_PLAN] Не удалось пометить плиту: КП #{kp_id}, {plate_name} x{qty_to_mark}")
        
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
        
        # Выходим в меню планирования производства
        await state.clear()
        await callback.message.answer(
            "📋 Планирование производства плит\n\n"
            "Выберите действие:",
            reply_markup=production_menu_kb()
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
