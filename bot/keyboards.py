"""Клавиатуры бота"""
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton


def main_menu_kb() -> ReplyKeyboardMarkup:
    """Главное меню бота"""
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Получить КП")],
            [KeyboardButton(text="Оптимизация резов")],
            [KeyboardButton(text="Коммерческое предложение PDF")],
            [KeyboardButton(text="Сравнение результатов")],
        ],
        resize_keyboard=True
    )

