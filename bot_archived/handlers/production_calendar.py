"""Просмотр календарного плана производства"""
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import CallbackQuery
from aiogram.fsm.context import FSMContext

from ..keyboards import (
    production_menu_kb, calendar_days_kb, plans_list_kb
)
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import (
    load_plans_metadata, get_active_plan,
    get_global_days_info, get_all_tracks_from_plan,
    convert_lookup_keys_to_tuples, get_global_calendar_info
)

router = Router()


@router.callback_query(F.data == "view_calendar_plan")
async def view_calendar_plan(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки 'Календарный план'.
    
    НОВАЯ ЛОГИКА:
    - Загружает ВСЕ планы и объединяет их в единый календарь
    - Показывает календарь с датами и СУММАРНОЙ загруженностью (3/5)
    - Выполненные дни отмечены галочкой ✅
    - По клику на дату — документы этого дня из ВСЕХ планов
    """
    # Загружаем глобальную информацию о календаре (из всех планов)
    global_calendar = get_global_calendar_info()
    
    if not global_calendar:
        # Проверяем, есть ли вообще планы
        metadata = load_plans_metadata()
        plans = metadata.get('plans', [])
        
        if plans:
            # Планы есть, но что-то пошло не так — показываем список
            await callback.message.answer(
                "❌ Не удалось загрузить календарь.\n\n"
                "Попробуйте выбрать план из списка:",
                reply_markup=plans_list_kb(plans, metadata.get('active_plan_id'))
            )
        else:
            # Планов нет вообще
            await callback.message.answer(
                "📭 Нет сохранённых планов\n\n"
                "Как создать план:\n"
                "1️⃣ Нажмите «🚀 Начать планирование»\n"
                "2️⃣ Введите дату начала и количество дорожек\n"
                "3️⃣ Выберите КП для производства\n"
                "4️⃣ Нажмите «💾 Сохранить план»\n\n"
                "💡 Только после сохранения план появится здесь!",
                reply_markup=production_menu_kb()
            )
        await callback.answer()
        return
    
    # Извлекаем данные из глобального календаря
    total_days = global_calendar['total_days']
    start_date = global_calendar['start_date']
    days_info = global_calendar['days_info']
    completed_days = global_calendar['completed_days']
    plans_count = global_calendar['plans_count']
    total_tracks_count = global_calendar['tracks_count']
    
    # Загружаем данные в state для работы с календарем
    await state.update_data(
        total_days=total_days,
        plan_start_date=start_date,
        completed_days=completed_days,
        from_saved_plan=True,
        days_info=days_info,
        is_global_calendar=True,  # Флаг для обработки кликов на даты
        tracks_count=5  # Максимум дорожек
    )
    
    # Считаем статистику
    completed_count = len(completed_days)
    remaining_count = total_days - completed_count
    
    # Формируем сообщение
    status_text = (
        f"📅 Календарный план производства\n\n"
        f"📊 Статистика (все планы):\n"
        f"  • Планов: {plans_count}\n"
        f"  • Всего дней: {total_days}\n"
        f"  • Выполнено: {completed_count} ✅\n"
        f"  • Осталось: {remaining_count}\n"
        f"  • Всего дорожек: {total_tracks_count}\n\n"
        f"Выберите день для просмотра документов:\n"
        f"(Формат: дата занято/максимум)"
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


@router.callback_query(F.data == "back_to_calendar")
async def back_to_calendar(callback: CallbackQuery, state: FSMContext):
    """
    Возвращает к календарному плану.
    
    Простыми словами:
    - Показывает календарь с датами снова
    - Используется для возврата из меню выбора документов
    - Берёт данные из памяти бота (state)
    """
    
    # Получаем данные для отображения календаря
    data = await state.get_data()
    total_days = data.get('total_days', 1)
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = data.get('days_info', {})
    from_saved_plan = data.get('from_saved_plan', False)
    
    await callback.message.answer(
        "📅 Выберите день для просмотра:",
        reply_markup=calendar_days_kb(
            total_days, 
            plan_start_date, 
            completed_days,
            days_info,
            show_save_button=not from_saved_plan
        )
    )
    
    await callback.answer()
