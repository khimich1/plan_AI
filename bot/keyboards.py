"""Клавиатуры бота"""
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton


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

