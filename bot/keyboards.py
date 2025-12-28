"""Клавиатуры бота"""
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton


def main_menu_kb() -> ReplyKeyboardMarkup:
    """Главное меню бота"""
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Получить КП")],
            [KeyboardButton(text="Коммерческое предложение PDF")],
            [KeyboardButton(text="Планирование производства")],
            [KeyboardButton(text="Сравнение результатов")],
        ],
        resize_keyboard=True
    )


def conditions_choice_kb() -> InlineKeyboardMarkup:
    """Экранная клавиатура выбора условий поставки и оплаты"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="По умолчанию", callback_data="conditions_default")],
            [InlineKeyboardButton(text="Добавить условие", callback_data="conditions_custom")],
        ]
    )


def save_to_db_kb() -> InlineKeyboardMarkup:
    """Клавиатура для сохранения КП в БД"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💾 Сохранить в БД", callback_data="save_kp_to_db")],
            [InlineKeyboardButton(text="❌ Не сохранять", callback_data="skip_save_kp")],
        ]
    )


def production_days_kb(total_days: int) -> InlineKeyboardMarkup:
    """
    Создает клавиатуру с кнопками для выбора дня производства
    
    Args:
        total_days: общее количество дней работы
        
    Returns:
        InlineKeyboardMarkup с кнопками "День 1", "День 2" и т.д.
    """
    buttons = []
    
    # Создаем кнопки по 3 в ряд для компактности
    row = []
    for day in range(1, total_days + 1):
        row.append(InlineKeyboardButton(
            text=f"📅 День {day}",
            callback_data=f"production_day_{day}"
        ))
        
        # Каждые 3 кнопки - новый ряд
        if len(row) == 3:
            buttons.append(row)
            row = []
    
    # Добавляем оставшиеся кнопки
    if row:
        buttons.append(row)
    
    # Кнопка "Все дни сразу" (если дней больше 1)
    if total_days > 1:
        buttons.append([
            InlineKeyboardButton(
                text="📦 Все дни сразу",
                callback_data="production_all_days"
            )
        ])
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)
