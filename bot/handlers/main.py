"""Основные команды бота: /start, /help, /stats"""
import os
from aiogram import Router
from aiogram.types import Message
from aiogram.filters import Command

from ..keyboards import main_menu_kb
from ..bot_config import OUTPUTS_DIR_STR

router = Router()


@router.message(Command("start"))
async def cmd_start(message: Message):
    """Обработчик команды /start"""
    await message.answer(
        "👋 Привет! Я бот для расчёта и визуализации дорожек ПБ.\n\n"
        "🔧 Что я умею:\n"
        "• Строить планы раскладки плит\n"
        "• Рассчитывать стоимость и отходы\n"
        "• Оптимизировать раскрой (экономия до 40%)\n"
        "• Экспортировать результаты в файлы\n\n"
        "Выберите действие кнопкой ниже или /help для справки",
        reply_markup=main_menu_kb()
    )


@router.message(Command("help"))
async def cmd_help(message: Message):
    """Обработчик команды /help"""
    help_text = """
📖 **Помощь по командам:**

🏗️ **Построить план** - создаёт визуализацию дорожки с расчётом стоимости

**Команды:**
• `/start` - главное меню
• `/build_plan` - построить план дорожки
• `/optimize` - оптимизация раскроя с экономией до 40%
• `/help` - эта справка
• `/stats` - статистика проекта

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
        await message.answer(f"❌ Ошибка: {str(e)}", parse_mode=None)

