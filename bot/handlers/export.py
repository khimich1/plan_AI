"""Обработчики экспорта заказов и истории"""
from aiogram import Router
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command

from ..keyboards import main_menu_kb

router = Router()


@router.message(Command("myorders"))
async def cmd_myorders(message: Message):
    """Показывает историю заказов пользователя"""
    try:
        import sqlite3
        from pathlib import Path
        
        # Проверяем существование БД
        db_path = Path(__file__).parent.parent / "pb.db"
        if not db_path.exists():
            await message.answer(
                "📋 База данных заказов не найдена.\n\n"
                "Создайте заказ через 'Коммерческое предложение PDF'",
                reply_markup=main_menu_kb()
            )
            return
        
        # TODO: Реализовать модуль domain.export для работы с историей заказов
        await message.answer(
            "⚠️ Функция истории заказов временно недоступна.\n\n"
            "Модуль domain.export находится в разработке.",
            reply_markup=main_menu_kb()
        )
        
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при получении истории заказов: {str(e)}",
            reply_markup=main_menu_kb()
        )


@router.message(Command("export"))
async def cmd_export(message: Message):
    """Экспортирует заказ в ZIP архив"""
    try:
        # Парсим ID заказа из команды /export_123
        command_parts = message.text.split('_')
        if len(command_parts) < 2:
            await message.answer(
                "❓ Укажите номер заказа: /export_123\n\n"
                "Посмотреть список заказов: /myorders",
                reply_markup=main_menu_kb()
            )
            return
        
        try:
            order_id = int(command_parts[1])
        except ValueError:
            await message.answer(
                "❌ Неверный формат номера заказа",
                reply_markup=main_menu_kb()
            )
            return
        
        # TODO: Реализовать модуль domain.export для создания архивов заказов
        await message.answer(
            f"⚠️ Экспорт заказа #{order_id} временно недоступен.\n\n"
            "Модуль domain.export находится в разработке.",
            reply_markup=main_menu_kb()
        )
        
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при экспорте: {str(e)}",
            reply_markup=main_menu_kb()
        )

