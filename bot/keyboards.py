"""Клавиатуры бота"""
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton


def main_menu_kb() -> ReplyKeyboardMarkup:
    """Главное меню бота"""
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Получить КП")],
            [KeyboardButton(text="Коммерческое предложение PDF")],
            [KeyboardButton(text="Планирование производства")],
            [KeyboardButton(text="Информация о ПБ в работе")],
            [KeyboardButton(text="Сравнение результатов")],
            [KeyboardButton(text="⚙️ Управление БД")],
        ],
        resize_keyboard=True
    )


def pb_info_kb() -> InlineKeyboardMarkup:
    """Клавиатура для выбора типа информации о плитах ПБ"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🏭 Плиты в производстве", callback_data="plates_in_production")],
            [InlineKeyboardButton(text="✅ Выполненные плиты", callback_data="completed_plates_export")],
        ]
    )


def conditions_choice_kb() -> InlineKeyboardMarkup:
    """Экранная клавиатура выбора условий поставки и оплаты"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="По умолчанию", callback_data="conditions_default")],
            [InlineKeyboardButton(text="Добавить условие", callback_data="conditions_custom")],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")],
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
    
    # Кнопка "Назад в меню"
    buttons.append([
        InlineKeyboardButton(
            text="◀️ Назад в меню",
            callback_data="cancel_process"
        )
    ])
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def production_day_actions_kb(day_number: int, total_days: int) -> InlineKeyboardMarkup:
    """
    Клавиатура действий после просмотра дня производства.
    Показывает кнопку "Выполнено" и кнопки для перехода на другие дни.
    
    Args:
        day_number: номер текущего дня
        total_days: общее количество дней
        
    Returns:
        InlineKeyboardMarkup с кнопками действий
    """
    buttons = []
    
    # Кнопка "День выполнен" для текущего дня
    buttons.append([
        InlineKeyboardButton(
            text="✅ День выполнен",
            callback_data=f"complete_day_{day_number}"
        )
    ])
    
    # Кнопки навигации по дням (по 3 в ряд)
    row = []
    for day in range(1, total_days + 1):
        row.append(InlineKeyboardButton(
            text=f"📅 День {day}",
            callback_data=f"production_day_{day}"
        ))
        if len(row) == 3:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def plates_completion_kb(plates_list: list, rejected_indices: set) -> InlineKeyboardMarkup:
    """
    Клавиатура для отметки бракованных плит при завершении дня.
    
    Простыми словами:
    - Показывает список плит с кнопками
    - Нажатие на плиту переключает её статус: брак/не брак
    - Внизу кнопки "Подтвердить" и "Отмена"
    
    Args:
        plates_list: список плит (словари с plate_name, qty и т.д.)
        rejected_indices: множество индексов бракованных плит
        
    Returns:
        InlineKeyboardMarkup с кнопками плит
    """
    buttons = []
    
    for idx, plate in enumerate(plates_list):
        # Если плита в браке — показываем ❌, иначе ✅
        is_rejected = idx in rejected_indices
        emoji = "❌ БРАК" if is_rejected else "✅"
        
        plate_name = plate.get('plate_name', f"Плита {idx+1}")
        qty = plate.get('qty', 1)
        
        # Обрезаем длинные названия
        if len(plate_name) > 25:
            plate_name = plate_name[:22] + "..."
        
        buttons.append([
            InlineKeyboardButton(
                text=f"{emoji} {plate_name} × {qty}",
                callback_data=f"toggle_reject_{idx}"
            )
        ])
    
    # Кнопки подтверждения
    buttons.append([
        InlineKeyboardButton(text="✅ Подтвердить", callback_data="confirm_completion"),
        InlineKeyboardButton(text="❌ Отмена", callback_data="cancel_completion")
    ])
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def db_management_kb() -> InlineKeyboardMarkup:
    """
    Клавиатура для управления базой данных.
    
    Показывает опции:
    - Очистка всех данных (КП, плиты, остатки)
    - Экспорт данных (будущая функция)
    - Статистика БД
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🗑️ Очистить все данные", callback_data="db_clear_all")],
            [InlineKeyboardButton(text="📊 Статистика БД", callback_data="db_stats")],
            [InlineKeyboardButton(text="◀️ Назад", callback_data="db_back_to_menu")],
        ]
    )


def db_clear_confirm_kb() -> InlineKeyboardMarkup:
    """
    Клавиатура подтверждения полной очистки БД.
    
    ВАЖНО: Это необратимая операция!
    Удаляет ВСЕ КП, плиты, выполненные плиты и остатки.
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="⚠️ ДА, УДАЛИТЬ ВСЁ", callback_data="db_clear_confirmed")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="db_clear_cancel")],
        ]
    )


def cancel_process_kb() -> InlineKeyboardMarkup:
    """Кнопка для отмены процесса и возврата в меню"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")]
        ]
    )