"""Главная точка входа для планирования производства плит"""
import logging

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import Message, CallbackQuery
from aiogram.fsm.context import FSMContext

from ..keyboards import main_menu_kb, production_menu_kb

router = Router()


@router.message(F.text == "Планирование производства")
async def btn_production_planning(message: Message, state: FSMContext):
    """
    Обработчик кнопки 'Планирование производства'.
    
    Показывает меню с вариантами:
    - Календарный план — просмотр активного плана с датами
    - Начать планирование — создание нового плана
    - Планы — просмотр всех сохранённых планов
    """
    await state.clear()  # Очищаем предыдущее состояние
    await message.answer(
        "📋 Планирование производства плит\n\n"
        "Выберите действие:",
        reply_markup=production_menu_kb()
    )


@router.callback_query(F.data == "cancel_process")
async def cancel_production_process(callback: CallbackQuery, state: FSMContext):
    """Отмена процесса планирования производства"""
    await state.clear()
    await callback.message.answer(
        "❌ Планирование производства отменено.\n"
        "Выберите действие:",
        reply_markup=main_menu_kb()
    )
    await callback.answer("Отменено")
