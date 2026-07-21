"""Управление рабочим календарем для планирования производства."""

from __future__ import annotations

import json
import logging
from datetime import datetime

from aiogram import F, Router
from aiogram.fsm.context import FSMContext
from aiogram.types import CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message

from core.work_calendar import CALENDAR_PATH
from ..keyboards import cancel_process_kb, production_menu_kb
from ..states import WorkCalendarStates

logger = logging.getLogger(__name__)
router = Router()


def _parse_date_input(value: str) -> datetime | None:
    """Парсит дату в формате DD.MM.YYYY или YYYY-MM-DD."""
    raw = (value or "").strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def _load_calendar_json() -> dict:
    """Загружает JSON календаря с безопасным fallback."""
    if not CALENDAR_PATH.exists():
        return {"extra_holidays": [], "extra_workdays": []}
    try:
        with open(CALENDAR_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"extra_holidays": [], "extra_workdays": []}
        return {
            "extra_holidays": list(data.get("extra_holidays", [])),
            "extra_workdays": list(data.get("extra_workdays", [])),
        }
    except Exception as exc:
        logger.exception("Ошибка загрузки work_calendar.json: %s", exc)
        return {"extra_holidays": [], "extra_workdays": []}


def _save_calendar_json(data: dict) -> bool:
    """Сохраняет JSON календаря."""
    try:
        CALENDAR_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CALENDAR_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as exc:
        logger.exception("Ошибка сохранения work_calendar.json: %s", exc)
        return False


def _work_calendar_menu_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📋 Список праздников", callback_data="wc_list_holidays")],
            [InlineKeyboardButton(text="➕ Добавить праздник", callback_data="wc_add_holiday")],
            [InlineKeyboardButton(text="➖ Удалить праздник", callback_data="wc_remove_holiday")],
            [InlineKeyboardButton(text="🛠️ Добавить рабочий день", callback_data="wc_add_workday")],
            [InlineKeyboardButton(text="◀️ Назад", callback_data="wc_back_to_production_menu")],
        ]
    )


@router.callback_query(F.data == "manage_work_calendar")
async def manage_work_calendar(callback: CallbackQuery, state: FSMContext):
    """Открывает меню управления производственным календарем."""
    await state.clear()
    await callback.message.answer(
        "🗓️ Производственный календарь\n\n"
        "Здесь можно добавлять/удалять праздничные дни и задавать рабочие переносы.",
        reply_markup=_work_calendar_menu_kb(),
    )
    await callback.answer()


@router.callback_query(F.data == "wc_list_holidays")
async def list_holidays(callback: CallbackQuery):
    """Показывает список праздников и дополнительных рабочих дней."""
    data = _load_calendar_json()
    holidays = sorted(set(data.get("extra_holidays", [])))
    workdays = sorted(set(data.get("extra_workdays", [])))

    lines = ["📋 Настройки рабочего календаря\n"]
    lines.append("Праздники / нерабочие дни:")
    if holidays:
        lines.extend([f"  • {d}" for d in holidays])
    else:
        lines.append("  • не заданы")

    lines.append("\nДополнительные рабочие дни:")
    if workdays:
        lines.extend([f"  • {d}" for d in workdays])
    else:
        lines.append("  • не заданы")

    await callback.message.answer("\n".join(lines), reply_markup=_work_calendar_menu_kb())
    await callback.answer()


@router.callback_query(F.data == "wc_add_holiday")
async def start_add_holiday(callback: CallbackQuery, state: FSMContext):
    """Запрашивает дату для добавления праздника."""
    await state.set_state(WorkCalendarStates.waiting_holiday_date)
    await callback.message.answer(
        "Введите дату праздника:\n"
        "• 31.12.2026\n"
        "• 2026-12-31",
        reply_markup=cancel_process_kb(),
    )
    await callback.answer()


@router.message(WorkCalendarStates.waiting_holiday_date)
async def save_holiday_date(message: Message, state: FSMContext):
    """Сохраняет праздничный день в календарь."""
    parsed = _parse_date_input(message.text or "")
    if not parsed:
        await message.answer("❌ Неверный формат даты. Используйте DD.MM.YYYY или YYYY-MM-DD.")
        return

    iso_day = parsed.strftime("%Y-%m-%d")
    data = _load_calendar_json()
    holidays = set(data.get("extra_holidays", []))
    workdays = set(data.get("extra_workdays", []))

    holidays.add(iso_day)
    # Если день помечен как праздник, убираем его из extra_workdays
    workdays.discard(iso_day)

    data["extra_holidays"] = sorted(holidays)
    data["extra_workdays"] = sorted(workdays)
    if not _save_calendar_json(data):
        await message.answer("❌ Не удалось сохранить календарь. Проверьте логи.")
        return

    await state.clear()
    await message.answer(
        f"✅ Праздничный день {iso_day} добавлен.",
        reply_markup=_work_calendar_menu_kb(),
    )


@router.callback_query(F.data == "wc_remove_holiday")
async def show_holidays_for_delete(callback: CallbackQuery):
    """Показывает список праздников для удаления."""
    data = _load_calendar_json()
    holidays = sorted(set(data.get("extra_holidays", [])))

    if not holidays:
        await callback.message.answer("Список праздников пуст.", reply_markup=_work_calendar_menu_kb())
        await callback.answer()
        return

    buttons = []
    for holiday in holidays[:25]:
        buttons.append([
            InlineKeyboardButton(
                text=f"🗑️ {holiday}",
                callback_data=f"wc_del_holiday_{holiday}",
            )
        ])
    buttons.append([InlineKeyboardButton(text="◀️ Назад", callback_data="manage_work_calendar")])

    await callback.message.answer(
        "Выберите праздник для удаления:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("wc_del_holiday_"))
async def delete_holiday(callback: CallbackQuery):
    """Удаляет выбранный праздник из календаря."""
    iso_day = callback.data.replace("wc_del_holiday_", "", 1)
    data = _load_calendar_json()
    holidays = set(data.get("extra_holidays", []))

    if iso_day not in holidays:
        await callback.answer("Дата уже удалена", show_alert=False)
        return

    holidays.remove(iso_day)
    data["extra_holidays"] = sorted(holidays)
    if not _save_calendar_json(data):
        await callback.message.answer("❌ Не удалось сохранить календарь. Проверьте логи.")
        await callback.answer()
        return

    await callback.message.answer(f"✅ Праздничный день {iso_day} удален.", reply_markup=_work_calendar_menu_kb())
    await callback.answer()


@router.callback_query(F.data == "wc_add_workday")
async def start_add_workday(callback: CallbackQuery, state: FSMContext):
    """Запрашивает дату дополнительного рабочего дня."""
    await state.set_state(WorkCalendarStates.waiting_workday_date)
    await callback.message.answer(
        "Введите дату дополнительного рабочего дня:\n"
        "• 10.01.2026\n"
        "• 2026-01-10",
        reply_markup=cancel_process_kb(),
    )
    await callback.answer()


@router.message(WorkCalendarStates.waiting_workday_date)
async def save_workday_date(message: Message, state: FSMContext):
    """Сохраняет дополнительный рабочий день."""
    parsed = _parse_date_input(message.text or "")
    if not parsed:
        await message.answer("❌ Неверный формат даты. Используйте DD.MM.YYYY или YYYY-MM-DD.")
        return

    iso_day = parsed.strftime("%Y-%m-%d")
    data = _load_calendar_json()
    holidays = set(data.get("extra_holidays", []))
    workdays = set(data.get("extra_workdays", []))

    workdays.add(iso_day)
    # Рабочий перенос имеет приоритет, убираем из праздников.
    holidays.discard(iso_day)

    data["extra_holidays"] = sorted(holidays)
    data["extra_workdays"] = sorted(workdays)
    if not _save_calendar_json(data):
        await message.answer("❌ Не удалось сохранить календарь. Проверьте логи.")
        return

    await state.clear()
    await message.answer(
        f"✅ Дополнительный рабочий день {iso_day} сохранен.",
        reply_markup=_work_calendar_menu_kb(),
    )


@router.callback_query(F.data == "wc_back_to_production_menu")
async def back_to_production_menu(callback: CallbackQuery, state: FSMContext):
    """Возвращает пользователя в меню планирования производства."""
    await state.clear()
    await callback.message.answer(
        "📋 Планирование производства\n\nВыберите действие:",
        reply_markup=production_menu_kb(),
    )
    await callback.answer()
