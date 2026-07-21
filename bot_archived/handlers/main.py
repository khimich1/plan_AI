"""Основные команды бота: /start, /help, /stats"""
import os
import logging
from aiogram import Router
from aiogram.types import Message
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext

from bot.security.users import BotUser

from ..keyboards import main_menu_kb
from ..bot_config import OUTPUTS_DIR_STR

router = Router()
logger = logging.getLogger(__name__)


@router.message(Command("start"))
async def cmd_start(message: Message, bot_user: BotUser):
    """Обработчик команды /start"""
    await message.answer(
        "👋 Привет! Я бот для расчёта и визуализации дорожек ПБ.\n\n"
        "🔧 Что я умею:\n"
        "• Строить планы раскладки плит\n"
        "• Рассчитывать стоимость и отходы\n"
        "• Оптимизировать раскрой (экономия до 40%)\n"
        "• Экспортировать результаты в файлы\n\n"
        "Выберите действие кнопкой ниже или /help для справки",
        reply_markup=main_menu_kb(bot_user.role),
    )


@router.message(Command("help"))
async def cmd_help(message: Message):
    """Обработчик команды /help"""
    help_text = """
📖 **Помощь по командам:**

🏗️ **Построить план** - создаёт визуализацию дорожки с расчётом стоимости

**Основные команды:**
• `/start` - главное меню
• `/build_plan` - построить план дорожки
• `/cancel` - отменить текущую операцию (если бот «завис» в диалоге)
• `/help` - эта справка
• `/stats` - статистика проекта

**Работа с КП:**
• `/list_kp` - список всех КП в базе данных
• `/delete_kp <номер>` - удалить КП по номеру

**Форматы файлов:**
• PNG - схема раскладки
• PDF - техническая документация  
• XLSX - ведомость и смета
• CSV - данные для импорта

💡 **Оптимизация резов:**
Использует каскадные продольные резы для минимизации отходов и экономии материала.
    """
    await message.answer(help_text, parse_mode="Markdown")


@router.message(Command("stats"))
async def cmd_stats(message: Message):
    """Обработчик команды /stats"""
    try:
        # Подсчитываем файлы в папке outputs
        files_count = len([f for f in os.listdir(OUTPUTS_DIR_STR) if f.endswith(('.png', '.pdf', '.xlsx'))])
        
        stats_text = f"""
📊 **Статистика проекта:**

📁 Файлов создано: {files_count}
📂 Папка результатов: `{OUTPUTS_DIR_STR}`

🔧 **Доступные функции:**
• Визуализация раскладки
• Расчёт стоимости материалов
• Экспорт в различные форматы

📈 **Последние результаты:**
• PNG схемы: {len([f for f in os.listdir(OUTPUTS_DIR_STR) if f.endswith('.png')])} шт
• PDF документы: {len([f for f in os.listdir(OUTPUTS_DIR_STR) if f.endswith('.pdf')])} шт
• Excel файлы: {len([f for f in os.listdir(OUTPUTS_DIR_STR) if f.endswith('.xlsx')])} шт
        """
        
        await message.answer(stats_text, parse_mode="Markdown")
        
    except Exception as e:
        logger.exception(f"Ошибка в /stats: {e}")
        await message.answer(
            "❌ Не удалось получить статистику.\n"
            "Попробуйте позже.",
            parse_mode=None
        )


@router.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext, bot_user: BotUser):
    """
    Универсальная отмена.

    Простыми словами:
    - если ты «застрял» в каком-то шаге (бот ждёт ввод),
      команда /cancel сбросит это состояние и вернёт в меню.
    """
    try:
        await state.clear()
    except Exception as e:
        logger.exception(f"Ошибка при /cancel: {e}")

    await message.answer(
        "❌ Операция отменена.\nВыберите действие:",
        reply_markup=main_menu_kb(bot_user.role),
    )

