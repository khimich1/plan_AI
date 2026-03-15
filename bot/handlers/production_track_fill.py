"""
Дозаполнение недозаполненных дорожек плана.
Вход из календаря одного плана: кнопка «Дозаполнить недозаполненные дорожки».
"""
import logging
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup
from aiogram.fsm.context import FSMContext

import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as cfg
from core import kp_db

from ..keyboards import calendar_days_kb
from ..states import ProductionStates

from .plan_manager import (
    load_plan,
    save_plan,
    update_plan_metadata,
    get_underfilled_tracks,
    get_global_days_info,
    load_plans_metadata,
    get_global_calendar_info,
)

router = Router()

MIN_TRACK_LENGTH = 96.0
MAX_TRACK_LENGTH = 101.0
PLATES_PAGE_SIZE = 8
DEFAULT_DB = str(PROJECT_ROOT / "plita.db")


def _plate_to_track_item(plate: dict) -> dict:
    """Преобразует запись плиты из БД в формат item дорожки плана."""
    length_m = float(plate.get('length_m', 0) or 0)
    width_m = float(plate.get('width_m', 1.2) or 1.2)
    load_class = plate.get('load_class', 800)
    load_code = cfg.normalize_load_code(load_class, default=8)
    # reinforcement для отображения; в плане часто 17.0 для 8п
    reinforcement = float(load_class) / 100.0 * 2.0 if load_class else 0
    kp_date = plate.get('execution_terms') or ''
    if isinstance(kp_date, str) and len(kp_date) > 10:
        try:
            # попытка формата DD.MM.YYYY
            datetime.strptime(kp_date[:10], '%d.%m.%Y')
            kp_date = kp_date[:10]
        except ValueError:
            kp_date = kp_date[:10] if len(kp_date) >= 10 else ''
    return {
        'length': round(length_m, 2),
        'mode': 'solid',
        'width': width_m,
        'load_code': load_code,
        'label': f"Плиты ПБ {plate.get('plate_name', '')}"[:80],
        'reinforcement': reinforcement,
        'is_separator': False,
        'kp_id': plate.get('kp_id'),
        'customer': plate.get('customer_name', ''),
        'kp_date': kp_date,
        'plate_name': plate.get('plate_name', ''),
    }


def _filter_plates_for_track(plates: list, load_code: int, max_reinforcement: float) -> list:
    """Оставляет плиты с подходящим load_code и reinforcement <= max_reinforcement."""
    out = []
    for p in plates:
        lc = cfg.normalize_load_code(p.get('load_class'), default=8)
        if lc != load_code:
            continue
        reinf = float(p.get('load_class', 0) or 0) / 100.0 * 2.0
        if max_reinforcement > 0 and reinf > max_reinforcement:
            continue
        out.append(p)
    return out


@router.callback_query(F.data == "fill_underfilled_tracks", ProductionStates.viewing_calendar)
async def handle_fill_underfilled_tracks(callback: CallbackQuery, state: FSMContext):
    """Показать список недозаполненных дорожек и кнопки выбора."""
    data = await state.get_data()
    is_global_calendar = data.get('is_global_calendar', False)

    if is_global_calendar:
        # Собираем недозаполненные дорожки по всем планам
        metadata = load_plans_metadata()
        plans_meta = metadata.get('plans', [])
        underfilled = []
        for plan_meta in plans_meta:
            plan_id = plan_meta.get('id')
            plan = load_plan(plan_id)
            if not plan:
                continue
            plan_name = plan.get('name', plan_id)
            for u in get_underfilled_tracks(plan, min_length=MIN_TRACK_LENGTH, max_length=MAX_TRACK_LENGTH):
                u = dict(u)
                u['plan_id'] = plan_id
                u['plan_name'] = plan_name
                underfilled.append(u)
    else:
        # Календарь одного плана
        plan_id = data.get('active_plan_id')
        if not plan_id:
            await callback.message.answer("❌ Не выбран активный план.")
            await callback.answer()
            return
        plan = load_plan(plan_id)
        if not plan:
            await callback.message.answer("❌ План не найден.")
            await callback.answer()
            return
        underfilled = get_underfilled_tracks(plan, min_length=MIN_TRACK_LENGTH, max_length=MAX_TRACK_LENGTH)
        for u in underfilled:
            u.setdefault('plan_id', plan_id)
            u.setdefault('plan_name', plan.get('name', plan_id))

    if not underfilled:
        await callback.message.answer(
            "✅ Недозаполненных дорожек не найдено (все дорожки не короче 96 м)."
        )
        await callback.answer()
        return

    lines = ["🛤️ Недозаполненные дорожки (длина < 96 м):\n"]
    buttons = []
    for i, u in enumerate(underfilled):
        date_key = u['date_key']
        day_number = u['day_number']
        track_idx = u['track_idx']
        track_length = u['track_length']
        free_space = u['free_space']
        plan_name = u.get('plan_name', '')
        try:
            dt = datetime.strptime(date_key, '%Y-%m-%d')
            date_str = dt.strftime('%d.%m')
        except Exception:
            date_str = date_key
        if is_global_calendar and plan_name:
            lines.append(
                f"{i + 1}) {plan_name} — День {day_number} ({date_str}), дор. {track_idx + 1} — {track_length} м, свободно {free_space} м"
            )
            cb = f"fill_track_{u['plan_id']}_{date_key}_{track_idx}"
            btn_text = f"{plan_name[:20]}… День {day_number} ({date_str}), дор. {track_idx + 1}" if len(plan_name) > 20 else f"{plan_name} — День {day_number} ({date_str}), дор. {track_idx + 1}"
        else:
            lines.append(
                f"{i + 1}) День {day_number} ({date_str}), дор. {track_idx + 1} — {track_length} м, свободно {free_space} м"
            )
            cb = f"fill_track_{date_key}_{track_idx}"
            btn_text = f"День {day_number} ({date_str}), дор. {track_idx + 1}"
        buttons.append([InlineKeyboardButton(text=btn_text, callback_data=cb)])

    buttons.append([InlineKeyboardButton(text="◀️ Отмена", callback_data="fill_cancel")])
    await state.update_data(
        fill_plan_id=data.get('active_plan_id') if not is_global_calendar else None,
        fill_underfilled_list=underfilled,
    )
    await state.set_state(ProductionStates.filling_track)
    await callback.message.answer(
        "\n".join(lines),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
    )
    await callback.answer()


@router.callback_query(F.data == "fill_cancel")
async def handle_fill_cancel(callback: CallbackQuery, state: FSMContext):
    """Отмена дозаполнения и возврат к календарю."""
    keys = ['fill_plan_id', 'fill_underfilled_list', 'fill_date_key', 'fill_track_idx',
            'fill_free_space', 'fill_max_reinforcement', 'fill_load_code', 'fill_selected_items',
            'fill_candidate_plates', 'fill_manual_page']
    data = await state.get_data()
    for k in keys:
        data.pop(k, None)
    await state.set_data(data)
    await state.set_state(ProductionStates.viewing_calendar)

    is_global_calendar = data.get('is_global_calendar', False)
    show_fill = is_global_calendar

    if is_global_calendar:
        global_calendar = get_global_calendar_info()
        if global_calendar:
            total_days = global_calendar['total_days']
            plan_start_date = global_calendar['start_date']
            completed_days = global_calendar['completed_days']
            days_info = global_calendar['days_info']
        else:
            total_days = data.get('total_days', 1)
            plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
            completed_days = data.get('completed_days', [])
            days_info = data.get('days_info', {})
    else:
        plan_id = data.get('active_plan_id')
        total_days = data.get('total_days', 1)
        plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
        completed_days = data.get('completed_days', [])
        plan = load_plan(plan_id) if plan_id else None
        days_info = get_global_days_info(plan) if plan else data.get('days_info', {})

    await callback.message.answer(
        "📅 Выберите день для просмотра:",
        reply_markup=calendar_days_kb(
            total_days,
            plan_start_date,
            completed_days,
            days_info,
            show_save_button=False,
            show_fill_underfilled=show_fill,
        ),
    )
    await callback.answer()


def _parse_fill_track_callback(data: str) -> tuple:
    """Парсит callback_data.
    Формат одного плана: fill_track_2026-02-20_1 -> (None, date_key, track_idx).
    Формат глобального: fill_track_plan_20260216_220508_2026-02-20_1 -> (plan_id, date_key, track_idx).
    Возвращает (plan_id или None, date_key, track_idx).
    """
    if not data.startswith("fill_track_"):
        return None, None, None
    rest = data.replace("fill_track_", "", 1)
    parts = rest.rsplit("_", 2)
    try:
        if len(parts) == 3 and "-" in parts[1]:  # date_key содержит дефис YYYY-MM-DD
            plan_id, date_key, track_idx = parts[0], parts[1], int(parts[2])
            return plan_id, date_key, track_idx
        if len(parts) == 2:
            date_key, track_idx = parts[0], int(parts[1])
            return None, date_key, track_idx
    except (ValueError, IndexError):
        pass
    return None, None, None


@router.callback_query(F.data.startswith("fill_track_"), ProductionStates.filling_track)
async def handle_fill_track_choice(callback: CallbackQuery, state: FSMContext):
    """Выбрана дорожка: предложить авто или вручную."""
    plan_id_from_cb, date_key, track_idx = _parse_fill_track_callback(callback.data)
    if date_key is None:
        await callback.answer("Ошибка параметров")
        return

    data = await state.get_data()
    underfilled = data.get('fill_underfilled_list', [])
    entry = None
    for u in underfilled:
        if u['date_key'] != date_key or u['track_idx'] != track_idx:
            continue
        if plan_id_from_cb is not None:
            if u.get('plan_id') == plan_id_from_cb:
                entry = u
                break
        else:
            entry = u
            break
    if not entry:
        await callback.message.answer("❌ Дорожка не найдена.")
        await callback.answer()
        return

    fill_plan_id = plan_id_from_cb if plan_id_from_cb is not None else data.get('fill_plan_id') or data.get('active_plan_id')
    await state.update_data(
        fill_plan_id=fill_plan_id,
        fill_date_key=date_key,
        fill_track_idx=track_idx,
        fill_free_space=entry['free_space'],
        fill_max_reinforcement=entry['max_reinforcement'],
        fill_load_code=entry['load_code'],
        fill_selected_items=[],
        fill_manual_page=0,
    )

    try:
        dt = datetime.strptime(date_key, '%Y-%m-%d')
        date_str = dt.strftime('%d.%m')
    except Exception:
        date_str = date_key

    msg = (
        f"Дорожка: День {entry['day_number']} ({date_str}), дор. {track_idx + 1}\n"
        f"Текущая длина: {entry['track_length']} м, свободно: {entry['free_space']} м.\n\n"
        "Как подобрать плиты?"
    )
    buttons = [
        [InlineKeyboardButton(text="🤖 Подобрать автоматически", callback_data="fill_auto_pick")],
        [InlineKeyboardButton(text="✋ Выбрать плиты вручную", callback_data="fill_manual_start")],
        [InlineKeyboardButton(text="◀️ Отмена", callback_data="fill_cancel")],
    ]
    await callback.message.answer(msg, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await callback.answer()


@router.callback_query(F.data == "fill_auto_pick", ProductionStates.filling_track)
async def handle_fill_auto_pick(callback: CallbackQuery, state: FSMContext):
    """Авто-подбор плит из «в производстве» и превью."""
    data = await state.get_data()
    free_space = data.get('fill_free_space', 0)
    max_reinforcement = data.get('fill_max_reinforcement', 0)
    load_code = data.get('fill_load_code', 8)

    plates = kp_db.get_all_plates_in_production(DEFAULT_DB)
    candidates = _filter_plates_for_track(plates, load_code, max_reinforcement)
    # Сортируем по длине по убыванию (жадное заполнение)
    candidates.sort(key=lambda p: float(p.get('length_m', 0) or 0), reverse=True)

    selected_items = []
    used_len = 0.0
    for p in candidates:
        length_m = float(p.get('length_m', 0) or 0)
        if length_m <= 0:
            continue
        if used_len + length_m > free_space or used_len + length_m > MAX_TRACK_LENGTH:
            continue
        item = _plate_to_track_item(p)
        item['_db_id'] = p.get('id')
        item['_kp_id'] = p.get('kp_id')
        item['_plate_name'] = p.get('plate_name')
        selected_items.append(item)
        used_len += length_m

    if not selected_items:
        await callback.message.answer(
            "Не найдено подходящих плит «в производстве» (та же нагрузка и армирование). Попробуйте выбрать вручную."
        )
        await callback.answer()
        return

    await state.update_data(fill_selected_items=selected_items)
    await _send_fill_preview(callback.message, state, selected_items)
    await callback.answer()


async def _send_fill_preview(message, state: FSMContext, selected_items: list):
    """Отправить превью добавленных плит и кнопки Подтвердить / Отмена."""
    data = await state.get_data()
    track_length = 0.0
    # Текущая длина дорожки из плана
    plan_id = data.get('fill_plan_id')
    date_key = data.get('fill_date_key')
    track_idx = data.get('fill_track_idx')
    plan = load_plan(plan_id)
    if plan and date_key and track_idx is not None:
        track = plan['days'].get(date_key, {}).get('tracks', [])[track_idx] if plan.get('days') else {}
        if track:
            track_length = track.get('length') or sum(
                float(i.get('length', 0) or 0) for i in track.get('items', [])
            )

    add_len = sum(float(i.get('length', 0) or 0) for i in selected_items)
    new_total = track_length + add_len

    lines = ["Превью добавления:\n"]
    for i, it in enumerate(selected_items[:15]):
        lines.append(f"  • {it.get('plate_name', '')} — {it.get('length', 0)} м")
    if len(selected_items) > 15:
        lines.append(f"  ... и ещё {len(selected_items) - 15} плит")
    lines.append(f"\nДобавляется: {add_len:.1f} м. Новая длина дорожки: {new_total:.1f} м.")

    buttons = [
        [InlineKeyboardButton(text="✅ Подтвердить", callback_data="fill_confirm_apply")],
        [InlineKeyboardButton(text="◀️ Отмена", callback_data="fill_cancel")],
    ]
    await message.answer(
        "\n".join(lines),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
    )


@router.callback_query(F.data == "fill_manual_start", ProductionStates.filling_track)
async def handle_fill_manual_start(callback: CallbackQuery, state: FSMContext):
    """Начать ручной выбор: показать первую страницу плит."""
    data = await state.get_data()
    load_code = data.get('fill_load_code', 8)
    max_reinforcement = data.get('fill_max_reinforcement', 0)

    plates = kp_db.get_all_plates_in_production(DEFAULT_DB)
    candidates = _filter_plates_for_track(plates, load_code, max_reinforcement)
    await state.update_data(fill_candidate_plates=candidates, fill_manual_page=0, fill_selected_items=[])
    await _send_manual_plates_page_async(callback.message, state, page=0)
    await state.set_state(ProductionStates.choosing_plates_to_fill)
    await callback.answer()


async def _send_manual_plates_page_async(message, state: FSMContext, page: int):
    data = await state.get_data()
    candidates = data.get('fill_candidate_plates', [])
    selected = data.get('fill_selected_items', [])
    free_space = data.get('fill_free_space', 0)
    track_length_cur = 0.0
    plan_id = data.get('fill_plan_id')
    date_key = data.get('fill_date_key')
    track_idx = data.get('fill_track_idx')
    plan = load_plan(plan_id)
    if plan and date_key is not None and track_idx is not None:
        day_data = plan.get('days', {}).get(date_key, {})
        tracks = day_data.get('tracks', [])
        if track_idx < len(tracks):
            t = tracks[track_idx]
            track_length_cur = t.get('length') or sum(float(i.get('length', 0) or 0) for i in t.get('items', []))

    selected_len = sum(float(i.get('length', 0) or 0) for i in selected)
    remaining = min(free_space, MAX_TRACK_LENGTH - track_length_cur - selected_len)
    if remaining <= 0:
        await message.answer("Свободного места больше нет. Нажмите «Готово».", reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Готово", callback_data="fill_manual_done")],
            [InlineKeyboardButton(text="◀️ Отмена", callback_data="fill_cancel")],
        ]))
        return

    start = page * PLATES_PAGE_SIZE
    chunk = candidates[start:start + PLATES_PAGE_SIZE]
    if not chunk and page == 0:
        await message.answer(
            "Нет подходящих плит «в производстве» (нагрузка и армирование дорожки).",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="◀️ Отмена", callback_data="fill_cancel")],
            ]),
        )
        return

    lines = [f"Выберите плиты (добавлено: {selected_len:.1f} м, свободно: {remaining:.1f} м):\n"]
    buttons = []
    for p in chunk:
        length_m = float(p.get('length_m', 0) or 0)
        name = p.get('plate_name', '')[:40]
        bid = p.get('id')
        if bid is None:
            continue
        buttons.append([
            InlineKeyboardButton(
                text=f"+ {name} ({length_m:.2f} м)",
                callback_data=f"fill_add_plate_{bid}",
            )
        ])

    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"fill_plate_page_{page - 1}"))
    if start + len(chunk) < len(candidates):
        nav.append(InlineKeyboardButton(text="Вперёд ➡️", callback_data=f"fill_plate_page_{page + 1}"))
    if nav:
        buttons.append(nav)
    buttons.append([InlineKeyboardButton(text="✅ Готово", callback_data="fill_manual_done")])
    buttons.append([InlineKeyboardButton(text="◀️ Отмена", callback_data="fill_cancel")])
    await message.answer("\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))


@router.callback_query(F.data.startswith("fill_plate_page_"), ProductionStates.choosing_plates_to_fill)
async def handle_fill_plate_page(callback: CallbackQuery, state: FSMContext):
    """Переключение страницы списка плит."""
    try:
        page = int(callback.data.replace("fill_plate_page_", ""))
    except ValueError:
        await callback.answer()
        return
    await state.update_data(fill_manual_page=page)
    await _send_manual_plates_page_async(callback.message, state, page)
    await callback.answer()


@router.callback_query(F.data.startswith("fill_add_plate_"), ProductionStates.choosing_plates_to_fill)
async def handle_fill_add_plate(callback: CallbackQuery, state: FSMContext):
    """Добавить одну плиту в выбранные (по id из БД)."""
    try:
        plate_id = int(callback.data.replace("fill_add_plate_", ""))
    except ValueError:
        await callback.answer()
        return

    data = await state.get_data()
    candidates = data.get('fill_candidate_plates', [])
    selected = list(data.get('fill_selected_items', []))
    free_space = data.get('fill_free_space', 0)
    plan_id = data.get('fill_plan_id')
    date_key = data.get('fill_date_key')
    track_idx = data.get('fill_track_idx')
    plan = load_plan(plan_id)
    track_length_cur = 0.0
    if plan and date_key is not None and track_idx is not None:
        day_data = plan.get('days', {}).get(date_key, {})
        tracks = day_data.get('tracks', [])
        if track_idx < len(tracks):
            t = tracks[track_idx]
            track_length_cur = t.get('length') or sum(float(i.get('length', 0) or 0) for i in t.get('items', []))

    plate = next((p for p in candidates if p.get('id') == plate_id), None)
    if not plate:
        await callback.answer("Плита не найдена")
        return

    length_m = float(plate.get('length_m', 0) or 0)
    selected_len = sum(float(i.get('length', 0) or 0) for i in selected)
    remaining = min(free_space, MAX_TRACK_LENGTH - track_length_cur - selected_len)
    if length_m > remaining:
        await callback.answer(f"Не влезает: свободно {remaining:.1f} м")
        return

    item = _plate_to_track_item(plate)
    item['_db_id'] = plate.get('id')
    item['_kp_id'] = plate.get('kp_id')
    item['_plate_name'] = plate.get('plate_name')
    selected.append(item)
    await state.update_data(fill_selected_items=selected)
    await callback.answer(f"Добавлено: {plate.get('plate_name', '')} ({length_m:.2f} м)")
    # Обновить текущую страницу
    page = data.get('fill_manual_page', 0)
    await _send_manual_plates_page_async(callback.message, state, page)


@router.callback_query(F.data == "fill_manual_done", ProductionStates.choosing_plates_to_fill)
async def handle_fill_manual_done(callback: CallbackQuery, state: FSMContext):
    """Готово с ручным выбором: превью и подтверждение."""
    data = await state.get_data()
    selected = data.get('fill_selected_items', [])
    await state.set_state(ProductionStates.filling_track)
    if not selected:
        await callback.message.answer("Не выбрано ни одной плиты. Выберите плиты или отмените.")
        await callback.answer()
        return
    await _send_fill_preview(callback.message, state, selected)
    await callback.answer()


@router.callback_query(F.data == "fill_confirm_apply")
async def handle_fill_confirm_apply(callback: CallbackQuery, state: FSMContext):
    """Применить дозаполнение: обновить план, сохранить, пометить плиты в БД."""
    data = await state.get_data()
    plan_id = data.get('fill_plan_id')
    date_key = data.get('fill_date_key')
    track_idx = data.get('fill_track_idx')
    selected = data.get('fill_selected_items', [])

    if not plan_id or not date_key or track_idx is None or not selected:
        await callback.message.answer("❌ Данные для сохранения не найдены.")
        await callback.answer()
        return

    plan = load_plan(plan_id)
    if not plan or date_key not in plan.get('days', {}):
        await callback.message.answer("❌ План или день не найден.")
        await callback.answer()
        return

    track = plan['days'][date_key]['tracks'][track_idx]
    items = list(track.get('items', []))
    # Убираем служебные поля из копий перед добавлением в план
    for it in selected:
        new_item = {k: v for k, v in it.items() if not k.startswith('_')}
        items.append(new_item)

    track['items'] = items
    track['length'] = round(sum(float(i.get('length', 0) or 0) for i in items), 2)
    track['max_reinforcement'] = max(float(i.get('reinforcement', 0) or 0) for i in items) if items else 0

    save_plan(plan)
    update_plan_metadata(plan)

    # Пометка плит «в плане» в БД
    marked = 0
    for it in selected:
        kp_id = it.get('_kp_id')
        plate_name = it.get('_plate_name')
        if kp_id and plate_name:
            result = kp_db.mark_plates_as_planned(
                kp_id=kp_id,
                plate_name=plate_name,
                qty_to_plan=1,
                plan_id=plan_id,
                db_path=DEFAULT_DB,
            )
            if result.get('success') and int(result.get('processed_count', 0) or 0) == 1:
                marked += 1
            else:
                logger.warning(
                    "[FILL_TRACK] Не удалось корректно пометить плиту при дозаполнении: "
                    "kp_id=%s, plate_name=%s, result=%s",
                    kp_id,
                    plate_name,
                    result,
                )

    logger.info(f"[FILL_TRACK] План {plan_id}, день {date_key}, дорожка {track_idx}: добавлено {len(selected)} плит, в БД помечено {marked}")

    # Очистка state и возврат к календарю
    keys = ['fill_plan_id', 'fill_underfilled_list', 'fill_date_key', 'fill_track_idx',
            'fill_free_space', 'fill_max_reinforcement', 'fill_load_code', 'fill_selected_items',
            'fill_candidate_plates', 'fill_manual_page']
    for k in keys:
        data.pop(k, None)
    await state.set_data(data)
    await state.set_state(ProductionStates.viewing_calendar)

    total_days = data.get('total_days', 1)
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = get_global_days_info(plan)

    await callback.message.answer(
        f"✅ Дорожка дозаполнена: добавлено {len(selected)} плит. В БД помечено как «в плане»: {marked}."
    )
    await callback.message.answer(
        "📅 Выберите день для просмотра:",
        reply_markup=calendar_days_kb(
            total_days,
            plan_start_date,
            completed_days,
            days_info,
            show_save_button=False,
            show_fill_underfilled=True,
        ),
    )
    await callback.answer()
