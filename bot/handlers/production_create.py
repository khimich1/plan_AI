"""Создание нового плана производства и фильтрация КП"""
import logging
import math
import sqlite3
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
from core.config_and_data import TRACK_LENGTH_M
from ..keyboards import (
    cancel_process_kb, production_filter_kb, main_menu_kb, tracks_choice_kb
)
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import load_plans_metadata, get_global_day_occupancy

router = Router()


@router.callback_query(F.data == "start_new_planning")
async def start_new_planning(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки 'Начать планирование'.
    
    Спрашивает дату начала плана.
    
    ВАЖНО: Сбрасываем active_plan_id, чтобы при сохранении 
    создавался НОВЫЙ план, а не добавлялись дорожки к существующему!
    """
    # Сбрасываем активный план — при сохранении будет создан НОВЫЙ план
    await state.update_data(active_plan_id=None)
    
    await state.set_state(ProductionStates.waiting_start_date)

    # Показываем, на какие дни уже есть планы и сколько в них дорожек
    occupancy = get_global_day_occupancy()
    if occupancy:
        sorted_dates = sorted(occupancy.keys())
        lines = ["📊 Дни с планами (дорожек в день):"]
        for date_key in sorted_dates:
            count = occupancy[date_key]
            try:
                dt = datetime.strptime(date_key, "%Y-%m-%d")
                label = dt.strftime("%d.%m")
            except ValueError:
                label = date_key
            lines.append(f"  • {label} — {count} дор.")
        occupancy_text = "\n".join(lines) + "\n\n"
    else:
        occupancy_text = "📊 Пока нет сохранённых планов по дням.\n\n"

    await callback.message.answer(
        "📅 Шаг 1 из 3: С какого числа начинается ваш план?\n\n"
        + occupancy_text
        + "Поддерживаемые форматы:\n"
        "• 22 (число текущего месяца)\n"
        "• 22.01.2026 (полная дата)\n"
        "• 2026-01-22 (ISO формат)\n\n"
        "Или оставьте пустым для сегодняшней даты:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.message(ProductionStates.waiting_start_date)
async def receive_start_date(message: Message, state: FSMContext):
    """
    Получаем дату начала плана.
    
    Простыми словами:
    - Парсим дату из разных форматов
    - Сохраняем в state
    - Переходим к вопросу о количестве дорожек
    """
    user_input = message.text.strip()
    
    # Если пустой ввод — используем сегодня
    if not user_input:
        start_date = datetime.now()
        date_description = "сегодня (" + start_date.strftime('%d.%m.%Y') + ")"
    else:
        start_date = None
        date_description = ""
        
        # Формат 1: Полная дата "ДД.ММ.ГГГГ"
        try:
            start_date = datetime.strptime(user_input, '%d.%m.%Y')
            date_description = start_date.strftime('%d.%m.%Y')
        except ValueError:
            pass
        
        # Формат 2: Полная дата "ГГГГ-ММ-ДД"
        if not start_date:
            try:
                start_date = datetime.strptime(user_input, '%Y-%m-%d')
                date_description = start_date.strftime('%d.%m.%Y')
            except ValueError:
                pass
        
        # Формат 3: Только число месяца
        if not start_date:
            try:
                date_number = int(user_input)
                
                if date_number < 1 or date_number > 31:
                    await message.answer(
                        "❌ Число должно быть от 1 до 31.\n"
                        "Попробуйте снова:",
                        reply_markup=cancel_process_kb()
                    )
                    return
                
                now = datetime.now()
                start_date = datetime(now.year, now.month, date_number)
                date_description = start_date.strftime('%d.%m.%Y')
            except ValueError:
                pass
        
        if not start_date:
            await message.answer(
                "❌ Неверный формат даты.\n\n"
                "Поддерживаемые форматы:\n"
                "• 22 (число месяца)\n"
                "• 22.01.2026 (полная дата)\n"
                "• 2026-01-22 (ISO формат)\n\n"
                "Попробуйте снова:",
                reply_markup=cancel_process_kb()
            )
            return
    
    # Сохраняем дату начала
    await state.update_data(
        plan_start_date=start_date.strftime('%Y-%m-%d'),
        plan_start_description=date_description
    )
    
    # Переходим к вопросу о дорожках
    await state.set_state(ProductionStates.waiting_tracks_count)
    await message.answer(
        f"✅ Дата начала плана: {date_description}\n\n"
        "📋 Шаг 2 из 3: Сколько дорожек нужно загрузить в день?\n"
        "(Выберите число или введите вручную)",
        reply_markup=tracks_choice_kb()
    )


@router.message(ProductionStates.waiting_tracks_count)
async def receive_tracks_count(message: Message, state: FSMContext):
    """Получаем количество дорожек (текстовый ввод)"""
    try:
        tracks_count = int(message.text.strip())
        
        if tracks_count <= 0 or tracks_count > 50:
            await message.answer(
                "❌ Количество дорожек должно быть от 1 до 50.\n"
                "Попробуйте снова:"
            )
            await message.answer(
                "Или выберите число на кнопках ниже:",
                reply_markup=tracks_choice_kb()
            )
            return
    except ValueError:
        await message.answer(
            "❌ Неверный формат. Введите целое число (например: 5):"
        )
        await message.answer(
            "Или выберите число на кнопках ниже:",
            reply_markup=tracks_choice_kb()
        )
        return
    
    await state.update_data(tracks_count=tracks_count)
    
    # Переходим к выбору способа фильтрации
    await state.set_state(ProductionStates.waiting_filter_method)
    await message.answer(
        f"✅ Дорожек: {tracks_count}\n\n"
        "📋 Шаг 3 из 3: Как выбрать плиты для производства?\n\n"
        "Выберите способ:",
        reply_markup=production_filter_kb()
    )


@router.callback_query(F.data.startswith("tracks_"), ProductionStates.waiting_tracks_count)
async def process_tracks_choice(callback: CallbackQuery, state: FSMContext):
    """Обработчик нажатия на кнопку с количеством дорожек"""
    tracks_count = int(callback.data.split("_")[1])

    await state.update_data(tracks_count=tracks_count)

    # Убираем кнопки у текущего сообщения
    await callback.message.edit_reply_markup(reply_markup=None)

    # Переходим к выбору способа фильтрации
    await state.set_state(ProductionStates.waiting_filter_method)
    await callback.message.answer(
        f"✅ Дорожек: {tracks_count}\n\n"
        "📋 Шаг 3 из 3: Как выбрать плиты для производства?\n\n"
        "Выберите способ:",
        reply_markup=production_filter_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "plan_step_back_to_1", ProductionStates.waiting_tracks_count)
async def plan_step_back_to_1(callback: CallbackQuery, state: FSMContext):
    """Шаг назад: с шага 2 (дорожки) на шаг 1 (дата начала)."""
    await state.set_state(ProductionStates.waiting_start_date)
    occupancy = get_global_day_occupancy()
    if occupancy:
        sorted_dates = sorted(occupancy.keys())
        lines = ["📊 Дни с планами (дорожек в день):"]
        for date_key in sorted_dates:
            count = occupancy[date_key]
            try:
                dt = datetime.strptime(date_key, "%Y-%m-%d")
                label = dt.strftime("%d.%m")
            except ValueError:
                label = date_key
            lines.append(f"  • {label} — {count} дор.")
        occupancy_text = "\n".join(lines) + "\n\n"
    else:
        occupancy_text = "📊 Пока нет сохранённых планов по дням.\n\n"
    try:
        await callback.message.edit_reply_markup(reply_markup=None)
    except Exception:
        pass
    await callback.message.answer(
        "📅 Шаг 1 из 3: С какого числа начинается ваш план?\n\n"
        + occupancy_text
        + "Поддерживаемые форматы:\n"
        "• 22 (число текущего месяца)\n"
        "• 22.01.2026 (полная дата)\n"
        "• 2026-01-22 (ISO формат)\n\n"
        "Или оставьте пустым для сегодняшней даты:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "plan_step_back_to_2", ProductionStates.waiting_filter_method)
async def plan_step_back_to_2(callback: CallbackQuery, state: FSMContext):
    """Шаг назад: с шага 3 (способ фильтрации) на шаг 2 (количество дорожек)."""
    await state.set_state(ProductionStates.waiting_tracks_count)
    data = await state.get_data()
    date_description = data.get("plan_start_description", "")
    try:
        await callback.message.edit_reply_markup(reply_markup=None)
    except Exception:
        pass
    await callback.message.answer(
        f"✅ Дата начала плана: {date_description}\n\n"
        "📋 Шаг 2 из 3: Сколько дорожек нужно загрузить в день?\n"
        "(Выберите число или введите вручную)",
        reply_markup=tracks_choice_kb()
    )
    await callback.answer()


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


def _build_kp_selection_message_and_kb(
    production_kp: list, selected_kp_ids: set, db_path: str, tracks_count: int = 1
):
    """Формирует текст и клавиатуру для экрана выбора КП (toggle + Плиты + Подтвердить + Назад)."""
    text = (
        "Нажмите на КП для выделения. Под списком — Подтвердить. "
        "Рядом с КП — кнопка «Плиты» для выбора плит по КП. "
        "В кнопке: первый % — выполнение КП, второй — плит в плане."
    )
    buttons = []
    for kp in production_kp:
        kp_id = kp['kp_id']
        plan_info = kp_db.get_kp_plates_in_plan_percentage(kp_id, db_path)
        in_plan_pct = plan_info['percentage']
        completion_info = kp_db.get_kp_completion_percentage(kp_id, db_path)
        completion_pct = completion_info['percentage']
        # Если КП полностью выполнен, в kp_plates не остаётся записей → 0% пл. Считаем это как 100% (всё было в плане и списано).
        in_plan_display = 100.0 if completion_pct >= 100 else in_plan_pct
        if in_plan_display >= 100:
            continue
        total_length = kp_db.get_kp_total_length(kp_id, db_path)
        estimated_tracks = max(1, round(total_length / TRACK_LENGTH_M))
        estimated_days = max(1, math.ceil(estimated_tracks / tracks_count))
        customer = kp.get('customer_name', 'Без имени')
        customer_short = customer[:12] + '…' if len(customer) > 12 else customer
        execution_terms = (kp.get('execution_terms') or '').strip()
        if len(execution_terms) > 10:
            execution_terms = execution_terms[:10].rstrip()
        prefix = "✓ " if kp_id in selected_kp_ids else ""
        btn_text = f"{prefix}КП №{kp_id} | {customer_short} | {completion_pct:.0f}% вып. · {in_plan_display:.0f}% пл. | ≈ {estimated_days} дн."
        if execution_terms:
            btn_text += f" | ⏰ {execution_terms}"
        buttons.append([
            InlineKeyboardButton(text=btn_text, callback_data=f"plan_kp_toggle_{kp_id}"),
        ])
        buttons.append([
            InlineKeyboardButton(text="▶ Плиты", callback_data=f"plan_kp_plates_{kp_id}"),
        ])
    if not buttons:
        return (
            "Все КП уже полностью в плане. Выберите другой способ фильтрации.",
            InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="◀️ Назад", callback_data="plan_kp_back")]
            ]),
        )
    buttons.append([
        InlineKeyboardButton(text="✅ Подтвердить", callback_data="plan_kp_confirm"),
        InlineKeyboardButton(text="◀️ Назад", callback_data="plan_kp_back"),
    ])
    return text, InlineKeyboardMarkup(inline_keyboard=buttons)


@router.callback_query(F.data == "filter_by_kp_buttons", ProductionStates.waiting_filter_method)
async def filter_by_kp_buttons(callback: CallbackQuery, state: FSMContext):
    """Выбор по КП: показываем список КП кнопками (мультивыбор + Плиты)."""
    await state.update_data(filter_method='kp', selected_kp_ids=[], kp_plate_ids={})
    await state.set_state(ProductionStates.selecting_kps)
    db_path = str(PROJECT_ROOT / "plita.db")
    all_kp = kp_db.get_all_kp_list(db_path)
    production_kp = all_kp.get('in_production', [])
    if not production_kp:
        await callback.message.answer(
            "❌ Нет КП в работе.\n\nВыберите другой способ или добавьте КП в производство.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="◀️ Назад", callback_data="plan_kp_back")]
            ])
        )
        await state.set_state(ProductionStates.waiting_filter_method)
        await callback.answer()
        return
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    selected_kp_ids = set()
    text, kb = _build_kp_selection_message_and_kb(production_kp, selected_kp_ids, db_path, tracks_count=tracks_count)
    try:
        await callback.message.edit_text(text, reply_markup=kb)
    except Exception:
        await callback.message.answer(text, reply_markup=kb)
    await callback.answer()


@router.callback_query(F.data.startswith("plan_kp_toggle_"), ProductionStates.selecting_kps)
async def plan_kp_toggle(callback: CallbackQuery, state: FSMContext):
    """Toggle выбора КП: добавляем/убираем из selected_kp_ids и обновляем сообщение."""
    kp_id = int(callback.data.replace("plan_kp_toggle_", ""))
    data = await state.get_data()
    selected = data.get('selected_kp_ids', [])
    selected_set = set(selected) if isinstance(selected, list) else set(selected)
    if kp_id in selected_set:
        selected_set.discard(kp_id)
    else:
        selected_set.add(kp_id)
    await state.update_data(selected_kp_ids=list(selected_set))
    db_path = str(PROJECT_ROOT / "plita.db")
    all_kp = kp_db.get_all_kp_list(db_path)
    production_kp = all_kp.get('in_production', [])
    if not production_kp:
        await callback.answer()
        return
    tracks_count = data.get('tracks_count', 1)
    text, kb = _build_kp_selection_message_and_kb(production_kp, selected_set, db_path, tracks_count=tracks_count)
    try:
        await callback.message.edit_text(text, reply_markup=kb)
    except Exception:
        pass
    await callback.answer()


@router.callback_query(F.data == "plan_kp_back", ProductionStates.selecting_kps)
async def plan_kp_back(callback: CallbackQuery, state: FSMContext):
    """Назад из списка КП — возврат к шагу 3 (выбор способа фильтрации)."""
    await state.set_state(ProductionStates.waiting_filter_method)
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    await callback.message.edit_text(
        f"✅ Дорожек: {tracks_count}\n\n"
        "📋 Шаг 3 из 3: Как выбрать плиты для производства?\n\n"
        "Выберите способ:",
        reply_markup=production_filter_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "plan_kp_confirm", ProductionStates.selecting_kps)
async def plan_kp_confirm(callback: CallbackQuery, state: FSMContext):
    """Подтвердить выбор КП и запустить планирование."""
    data = await state.get_data()
    selected = data.get('selected_kp_ids', [])
    selected_set = set(selected) if isinstance(selected, list) else set(selected)
    if not selected_set:
        await callback.answer("Выберите хотя бы одно КП", show_alert=True)
        return
    kp_plate_ids = data.get('kp_plate_ids', {})
    if not isinstance(kp_plate_ids, dict):
        kp_plate_ids = {}
    await state.update_data(kp_ids=list(selected_set), kp_plate_ids=kp_plate_ids)
    await callback.message.edit_reply_markup(reply_markup=None)
    await callback.message.answer(f"✅ Выбрано КП: {', '.join(map(str, sorted(selected_set)))}\n\n⏳ Загружаю плиты...")
    from .production_execution import load_and_plan_production
    await load_and_plan_production(callback.message, state)
    await callback.answer()


def _build_plates_selection_message_and_kb(kp_id: int, plates: list, selected_plate_ids: set):
    """Формирует текст и клавиатуру для выбора плит внутри КП."""
    text = f"КП №{kp_id}: выберите плиты для плана (нажмите для выделения)."
    buttons = []
    for p in plates:
        pid = p['id']
        name = (p.get('plate_name') or '')[:25]
        if len(p.get('plate_name') or '') > 25:
            name += '…'
        qty = p.get('qty', 1)
        prefix = "✓ " if pid in selected_plate_ids else ""
        btn_text = f"{prefix}{name} ×{qty}"
        buttons.append([
            InlineKeyboardButton(text=btn_text, callback_data=f"plan_plate_toggle_{pid}")
        ])
    buttons.append([
        InlineKeyboardButton(text="✅ Подтвердить", callback_data="plan_plates_confirm"),
        InlineKeyboardButton(text="◀️ Назад к списку КП", callback_data="plan_plates_back"),
    ])
    return text, InlineKeyboardMarkup(inline_keyboard=buttons)


@router.callback_query(F.data.startswith("plan_kp_plates_"), ProductionStates.selecting_kps)
async def plan_kp_plates_open(callback: CallbackQuery, state: FSMContext):
    """Открыть экран выбора плит для одного КП."""
    kp_id = int(callback.data.replace("plan_kp_plates_", ""))
    db_path = str(PROJECT_ROOT / "plita.db")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, plate_name, length_m, width_m, qty
        FROM kp_plates
        WHERE kp_id = ? AND status = 'в производстве'
        ORDER BY position_number, id
    """, (kp_id,))
    rows = cur.fetchall()
    conn.close()
    plates = [{'id': r[0], 'plate_name': r[1], 'length_m': r[2], 'width_m': r[3], 'qty': r[4]} for r in rows]
    if not plates:
        await callback.answer("В этом КП нет плит в производстве", show_alert=True)
        return
    data = await state.get_data()
    kp_plate_ids = data.get('kp_plate_ids', {}) or {}
    if not isinstance(kp_plate_ids, dict):
        kp_plate_ids = {}
    sk = str(kp_id)
    if sk in kp_plate_ids and kp_plate_ids[sk] is not None:
        selected_plate_ids = set(kp_plate_ids[sk])
    else:
        selected_plate_ids = set(p['id'] for p in plates)
    await state.set_state(ProductionStates.selecting_plates_in_kp)
    await state.update_data(
        current_kp_plates_kp_id=kp_id,
        selected_plate_ids_for_current_kp=list(selected_plate_ids),
    )
    text, kb = _build_plates_selection_message_and_kb(kp_id, plates, selected_plate_ids)
    try:
        await callback.message.edit_text(text, reply_markup=kb)
    except Exception:
        await callback.message.answer(text, reply_markup=kb)
    await callback.answer()


@router.callback_query(F.data.startswith("plan_plate_toggle_"), ProductionStates.selecting_plates_in_kp)
async def plan_plate_toggle(callback: CallbackQuery, state: FSMContext):
    """Toggle выбора плиты внутри КП."""
    plate_id = int(callback.data.replace("plan_plate_toggle_", ""))
    data = await state.get_data()
    kp_id = data.get('current_kp_plates_kp_id')
    selected = data.get('selected_plate_ids_for_current_kp', [])
    selected_set = set(selected) if isinstance(selected, list) else set(selected)
    if plate_id in selected_set:
        selected_set.discard(plate_id)
    else:
        selected_set.add(plate_id)
    await state.update_data(selected_plate_ids_for_current_kp=list(selected_set))
    db_path = str(PROJECT_ROOT / "plita.db")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, plate_name, length_m, width_m, qty
        FROM kp_plates
        WHERE kp_id = ? AND status = 'в производстве'
        ORDER BY position_number, id
    """, (kp_id,))
    rows = cur.fetchall()
    conn.close()
    plates = [{'id': r[0], 'plate_name': r[1], 'length_m': r[2], 'width_m': r[3], 'qty': r[4]} for r in rows]
    text, kb = _build_plates_selection_message_and_kb(kp_id, plates, selected_set)
    try:
        await callback.message.edit_text(text, reply_markup=kb)
    except Exception:
        pass
    await callback.answer()


async def _redraw_kp_list(callback: CallbackQuery, state: FSMContext):
    """Возврат на экран выбора КП и обновление сообщения."""
    await state.set_state(ProductionStates.selecting_kps)
    data = await state.get_data()
    selected = data.get('selected_kp_ids', [])
    selected_set = set(selected) if isinstance(selected, list) else set(selected)
    db_path = str(PROJECT_ROOT / "plita.db")
    all_kp = kp_db.get_all_kp_list(db_path)
    production_kp = all_kp.get('in_production', [])
    if not production_kp:
        await callback.message.edit_text(
            "❌ Нет КП в работе.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="◀️ Назад", callback_data="plan_kp_back")]
            ])
        )
        return
    tracks_count = data.get('tracks_count', 1)
    text, kb = _build_kp_selection_message_and_kb(production_kp, selected_set, db_path, tracks_count=tracks_count)
    try:
        await callback.message.edit_text(text, reply_markup=kb)
    except Exception:
        await callback.message.answer(text, reply_markup=kb)


@router.callback_query(F.data == "plan_plates_confirm", ProductionStates.selecting_plates_in_kp)
async def plan_plates_confirm(callback: CallbackQuery, state: FSMContext):
    """Подтвердить выбор плит по КП и вернуться к списку КП."""
    data = await state.get_data()
    kp_id = data.get('current_kp_plates_kp_id')
    selected = data.get('selected_plate_ids_for_current_kp', [])
    selected_set = set(selected) if isinstance(selected, list) else set(selected)
    kp_plate_ids = data.get('kp_plate_ids', {}) or {}
    if not isinstance(kp_plate_ids, dict):
        kp_plate_ids = {}
    kp_plate_ids[str(kp_id)] = list(selected_set) if selected_set else []
    await state.update_data(kp_plate_ids=kp_plate_ids)
    await _redraw_kp_list(callback, state)
    await callback.answer()


@router.callback_query(F.data == "plan_plates_back", ProductionStates.selecting_plates_in_kp)
async def plan_plates_back(callback: CallbackQuery, state: FSMContext):
    """Назад из выбора плит — вернуться к списку КП без сохранения."""
    await _redraw_kp_list(callback, state)
    await callback.answer()


@router.callback_query(F.data == "filter_all", ProductionStates.waiting_filter_method)
async def filter_all(callback: CallbackQuery, state: FSMContext):
    """Выбор всех КП в работе - сразу переходим к загрузке"""
    await state.update_data(filter_method='all')
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_execution import load_and_plan_production
    await load_and_plan_production(callback.message, state)
    await callback.answer()


@router.callback_query(F.data == "filter_by_customer", ProductionStates.waiting_filter_method)
async def filter_by_customer(callback: CallbackQuery, state: FSMContext):
    """Выбор по заказчику - показываем список заказчиков"""
    await state.update_data(filter_method='customer')
    
    # Загружаем список уникальных заказчиков из БД
    db_path = PROJECT_ROOT / "plita.db"
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
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_execution import load_and_plan_production
    await load_and_plan_production(callback.message, state)
    await callback.answer()


# === ОБРАБОТЧИКИ ВВОДА ДАТЫ (шаг 3 — по дате) ===

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
    
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_execution import load_and_plan_production
    await load_and_plan_production(message, state)
