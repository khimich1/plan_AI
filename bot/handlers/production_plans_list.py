"""Управление списком сохранённых планов производства"""
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import CallbackQuery
from aiogram.fsm.context import FSMContext

from ..keyboards import (
    production_menu_kb, calendar_days_kb, plans_list_kb, 
    plan_actions_kb, confirm_delete_plan_kb
)
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import (
    load_plans_metadata, load_plan, set_active_plan,
    get_active_plan_id, delete_plan as delete_plan_func,
    get_global_days_info, get_all_tracks_from_plan,
    convert_lookup_keys_to_tuples
)

router = Router()


@router.callback_query(F.data == "view_all_plans")
async def view_all_plans(callback: CallbackQuery, state: FSMContext):
    """
    Показывает список всех сохранённых планов.
    Позволяет выбрать план для просмотра, сделать активным или удалить.
    """
    # Загружаем метаданные планов
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])
    active_plan_id = metadata.get('active_plan_id')
    
    # Формируем сообщение
    if not plans:
        message_text = (
            "📋 Список планов\n\n"
            "У вас пока нет сохранённых планов.\n"
            "Нажмите «Создать новый план» для начала работы."
        )
    else:
        message_text = (
            f"📋 Список планов ({len(plans)} шт.)\n\n"
            f"Активный план отмечен звёздочкой ⭐\n"
            f"Выберите план для управления:"
        )
    
    await callback.message.answer(
        message_text,
        reply_markup=plans_list_kb(plans, active_plan_id)
    )
    await state.set_state(ProductionStates.viewing_plans_list)
    await callback.answer()


@router.callback_query(F.data == "no_plans_info")
async def no_plans_info(callback: CallbackQuery):
    """Обработчик нажатия на пустую кнопку 'Нет планов'."""
    await callback.answer("Создайте новый план для начала работы", show_alert=True)


@router.callback_query(F.data.startswith("select_plan_"))
async def select_plan(callback: CallbackQuery, state: FSMContext):
    """
    Загружает выбранный план и показывает меню действий с ним.
    """
    plan_id = callback.data.replace("select_plan_", "")
    
    # Загружаем план
    plan = load_plan(plan_id)
    if not plan:
        await callback.message.answer(
            "❌ План не найден.\n"
            "Возможно, он был удалён.",
            reply_markup=production_menu_kb()
        )
        await callback.answer()
        return
    
    # Получаем информацию об активном плане
    active_plan_id = get_active_plan_id()
    is_active = (plan_id == active_plan_id)
    
    # Подсчитываем статистику
    total_days = len(plan.get('days', {}))
    total_tracks = sum(
        day.get('saved_tracks_count', len(day.get('tracks', [])))
        for day in plan.get('days', {}).values()
    )
    completed_days = sum(
        1 for day in plan.get('days', {}).values()
        if day.get('completed', False)
    )
    
    # Форматируем дату начала
    start_date = plan.get('start_date', '')
    try:
        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
        start_date_str = start_dt.strftime('%d.%m.%Y')
    except:
        start_date_str = start_date
    
    # Формируем сообщение
    status_emoji = "⭐" if is_active else "📋"
    message_text = (
        f"{status_emoji} {plan.get('name', 'План')}\n\n"
        f"📅 Дата начала: {start_date_str}\n"
        f"📊 Дней в плане: {total_days}\n"
        f"🛤️ Всего дорожек: {total_tracks}\n"
        f"✅ Выполнено дней: {completed_days}\n"
        f"🚦 Статус: {'Активный' if is_active else 'Неактивный'}\n\n"
        f"Выберите действие:"
    )
    
    await callback.message.answer(
        message_text,
        reply_markup=plan_actions_kb(plan_id, is_active)
    )
    await callback.answer()


@router.callback_query(F.data == "create_new_plan")
async def create_new_plan(callback: CallbackQuery, state: FSMContext):
    """
    Начинает создание нового плана.
    Перенаправляет на start_new_planning из production_create.py.
    """
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_create import start_new_planning
    await start_new_planning(callback, state)


@router.callback_query(F.data.startswith("activate_plan_"))
async def activate_plan(callback: CallbackQuery, state: FSMContext):
    """
    Устанавливает выбранный план как активный.
    """
    plan_id = callback.data.replace("activate_plan_", "")
    
    # Проверяем существование плана
    plan = load_plan(plan_id)
    if not plan:
        await callback.message.answer(
            "❌ План не найден.",
            reply_markup=production_menu_kb()
        )
        await callback.answer()
        return
    
    # Устанавливаем как активный
    set_active_plan(plan_id)
    
    await callback.message.answer(
        f"⭐ План «{plan.get('name', plan_id)}» теперь активный!\n\n"
        f"Теперь при нажатии «Календарный план» будет открываться именно он."
    )
    
    # Показываем обновлённый список планов
    await view_all_plans(callback, state)


@router.callback_query(F.data.startswith("open_plan_calendar_"))
async def open_plan_calendar(callback: CallbackQuery, state: FSMContext):
    """
    Открывает календарь выбранного плана.
    """
    plan_id = callback.data.replace("open_plan_calendar_", "")
    
    # Загружаем план
    plan = load_plan(plan_id)
    if not plan:
        await callback.message.answer(
            "❌ План не найден.",
            reply_markup=production_menu_kb()
        )
        await callback.answer()
        return
    
    # Устанавливаем как активный (для удобства)
    set_active_plan(plan_id)
    
    # Получаем информацию о днях с ГЛОБАЛЬНОЙ загруженностью
    days_info = get_global_days_info(plan)
    total_days = len(plan.get('days', {}))
    start_date = plan.get('start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = plan.get('completed_days', [])
    
    # ИЗМЕНЕНО: НЕ загружаем all_tracks_list и детальные данные в state
    # Это заставит систему использовать мультиплан-логику при завершении дня
    # (загрузка данных из ВСЕХ планов на конкретную дату)
    
    # Загружаем только МЕТАДАННЫЕ в state
    await state.update_data(
        active_plan_id=plan_id,
        total_days=total_days,
        tracks_count=plan.get('tracks_count', 5),
        plan_start_date=start_date,
        completed_days=completed_days,
        from_saved_plan=True,
        days_info=days_info
    )
    
    # Подсчитываем статистику
    total_tracks = sum(d.get('occupied', 0) for d in days_info.values())
    completed_count = sum(1 for d in days_info.values() if d.get('completed', False))
    
    # Формируем сообщение
    status_text = (
        f"📅 {plan.get('name', 'План производства')}\n\n"
        f"📊 Статистика:\n"
        f"  • Всего дней: {total_days}\n"
        f"  • Выполнено: {completed_count} ✅\n"
        f"  • Осталось: {total_days - completed_count}\n"
        f"  • Всего дорожек: {total_tracks}\n\n"
        f"Выберите день для просмотра документов:"
    )
    
    await callback.message.answer(
        status_text,
        reply_markup=calendar_days_kb(
            total_days, 
            start_date, 
            completed_days, 
            days_info,
            show_save_button=False  # План уже сохранён
        )
    )
    
    await state.set_state(ProductionStates.viewing_calendar)
    await callback.answer()


@router.callback_query(F.data.startswith("delete_plan_"))
async def delete_plan_handler(callback: CallbackQuery, state: FSMContext):
    """
    Запрашивает подтверждение удаления плана.
    """
    plan_id = callback.data.replace("delete_plan_", "")
    
    # Загружаем план для отображения названия
    plan = load_plan(plan_id)
    if not plan:
        await callback.message.answer(
            "❌ План не найден.",
            reply_markup=production_menu_kb()
        )
        await callback.answer()
        return
    
    # Сохраняем ID плана для удаления
    await state.update_data(deleting_plan_id=plan_id)
    
    await callback.message.answer(
        f"⚠️ Удалить план «{plan.get('name', plan_id)}»?\n\n"
        f"Это действие нельзя отменить!\n"
        f"Все дорожки и данные плана будут удалены.",
        reply_markup=confirm_delete_plan_kb(plan_id)
    )
    await state.set_state(ProductionStates.confirming_plan_delete)
    await callback.answer()


@router.callback_query(F.data.startswith("confirm_delete_plan_"))
async def confirm_delete_plan(callback: CallbackQuery, state: FSMContext):
    """
    Подтверждение удаления плана.
    """
    plan_id = callback.data.replace("confirm_delete_plan_", "")
    
    # Загружаем план для отображения названия
    plan = load_plan(plan_id)
    plan_name = plan.get('name', plan_id) if plan else plan_id
    
    # Удаляем план
    success = delete_plan_func(plan_id)
    
    if success:
        await callback.message.answer(
            f"✅ План «{plan_name}» удалён."
        )
    else:
        await callback.message.answer(
            f"❌ Не удалось удалить план «{plan_name}»."
        )
    
    # Возвращаемся к списку планов
    await view_all_plans(callback, state)


@router.callback_query(F.data == "back_to_production_menu")
async def back_to_production_menu(callback: CallbackQuery, state: FSMContext):
    """
    Возврат в меню планирования производства.
    """
    await state.clear()
    await callback.message.answer(
        "📋 Планирование производства плит\n\n"
        "Выберите действие:",
        reply_markup=production_menu_kb()
    )
    await callback.answer()
