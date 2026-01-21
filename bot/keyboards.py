"""Клавиатуры бота"""
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton


def main_menu_kb() -> ReplyKeyboardMarkup:
    """Главное меню бота"""
    return ReplyKeyboardMarkup(
        keyboard=[
            # Первая строка - 2 кнопки рядом
            [
                KeyboardButton(text="📝 Создать КП"),
                KeyboardButton(text="📁 Архив")
            ],
            # Вторая строка - 2 кнопки рядом
            [
                KeyboardButton(text="Планирование производства"),
                KeyboardButton(text="Информация о ПБ в работе")
            ],
            # Третья строка - 2 кнопки рядом
            [
                KeyboardButton(text="Сравнение результатов"),
                KeyboardButton(text="⚙️ Управление БД")
            ],
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
            [InlineKeyboardButton(text="📦 В архив", callback_data="save_kp_to_archive")],
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
    
    # Кнопка "Диаграмма Ганта" вверху
    buttons.append([
        InlineKeyboardButton(
            text="📊 Диаграмма Ганта",
            callback_data="export_gantt"
        )
    ])
    
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
    
    # Кнопка "Актуальный план"
    buttons.append([
        InlineKeyboardButton(
            text="💾 Актуальный план",
            callback_data="save_current_plan"
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
    
    # Кнопка "Актуальный план"
    buttons.append([
        InlineKeyboardButton(
            text="💾 Актуальный план",
            callback_data="save_current_plan"
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


def production_day_completed_kb(current_day: int, total_days: int) -> InlineKeyboardMarkup:
    """
    Клавиатура, которая показывается ПОСЛЕ завершения дня.
    
    Кнопки:
    - "Перейти к следующему дню" — открывает оставшиеся дни (current_day+1 .. total_days)
    - "Назад в меню" — возвращает в главное меню
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="➡️ Перейти к следующему дню",
                    callback_data=f"production_next_days_{current_day}"
                )
            ],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")],
        ]
    )


def production_remaining_days_kb(from_day: int, total_days: int) -> InlineKeyboardMarkup:
    """
    Показывает ТОЛЬКО оставшиеся дни после from_day.
    Например: from_day=2, total_days=5 -> кнопки День 3, День 4, День 5.
    """
    buttons = []
    
    # Кнопка "Диаграмма Ганта" вверху (как в основном выборе дней)
    buttons.append([
        InlineKeyboardButton(
            text="📊 Диаграмма Ганта",
            callback_data="export_gantt"
        )
    ])
    
    # Дни (по 3 в ряд)
    row = []
    for day in range(from_day + 1, total_days + 1):
        row.append(InlineKeyboardButton(
            text=f"📅 День {day}",
            callback_data=f"production_day_{day}"
        ))
        if len(row) == 3:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    
    # "Актуальный план"
    buttons.append([
        InlineKeyboardButton(
            text="💾 Актуальный план",
            callback_data="save_current_plan"
        )
    ])
    
    # "Назад в меню"
    buttons.append([
        InlineKeyboardButton(
            text="◀️ Назад в меню",
            callback_data="cancel_process"
        )
    ])
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def plates_completion_kb(plates_by_track: list, rejected_indices: set) -> InlineKeyboardMarkup:
    """
    Клавиатура для отметки бракованных плит при завершении дня.
    
    Простыми словами:
    - Показывает список плит по дорожкам с кнопками
    - Каждая дорожка начинается с заголовка
    - Нажатие на плиту переключает её статус: брак/не брак
    - Внизу кнопки "Подтвердить" и "Отмена"
    
    Args:
        plates_by_track: список дорожек, каждая содержит track_number и список plates
        rejected_indices: множество кортежей (track_idx, plate_idx) для бракованных плит
        
    Returns:
        InlineKeyboardMarkup с кнопками плит
    """
    buttons = []
    
    for track_idx, track_data in enumerate(plates_by_track):
        track_number = track_data.get('track_number', track_idx + 1)
        plates = track_data.get('plates', [])
        
        if not plates:
            continue
        
        # Добавляем заголовок дорожки (кнопка без действия)
        buttons.append([
            InlineKeyboardButton(
                text=f"━━━━ Дорожка {track_number} ━━━━",
                callback_data=f"track_header_{track_idx}"
            )
        ])
        
        # Добавляем плиты этой дорожки
        for plate_idx, plate in enumerate(plates):
            # Если плита в браке — показываем ❌, иначе ✅
            is_rejected = (track_idx, plate_idx) in rejected_indices
            emoji = "❌" if is_rejected else "✅"
            
            plate_name = plate.get('plate_name', f"Плита {plate_idx+1}")
            qty = plate.get('qty', 1)
            kp_date = plate.get('kp_date', '')
            kp_id = plate.get('kp_id', '')
            
            # Форматируем дату коротко: "02.02.2026" -> "02.02"
            date_short = kp_date[:5] if kp_date and kp_date != 'неизвестно' else ''
            
            # Формируем информацию о КП
            if kp_id and date_short:
                kp_info = f"({date_short})"
            elif kp_id:
                kp_info = f"(КП{kp_id})"
            else:
                kp_info = ""
            
            # Обрезаем длинные названия (с учетом места для КП-инфо)
            max_name_len = 18 if kp_info else 25
            if len(plate_name) > max_name_len:
                plate_name = plate_name[:max_name_len-2] + ".."
            
            buttons.append([
                InlineKeyboardButton(
                    text=f"{emoji} {plate_name} {kp_info} × {qty}",
                    callback_data=f"toggle_reject_t{track_idx}_p{plate_idx}"
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


def managers_selection_kb(managers_list: list) -> InlineKeyboardMarkup:
    """
    Клавиатура для выбора менеджера из списка.
    
    Args:
        managers_list: список словарей с данными менеджеров из БД
        
    Returns:
        InlineKeyboardMarkup с кнопками менеджеров
    """
    buttons = []
    for manager in managers_list:
        buttons.append([
            InlineKeyboardButton(
                text=manager['fio'],
                callback_data=f"select_manager_{manager['id']}"
            )
        ])
    
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")
    ])
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def cancel_process_kb() -> InlineKeyboardMarkup:
    """Кнопка для отмены процесса и возврата в меню"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")]
        ]
    )


def archive_sections_kb() -> InlineKeyboardMarkup:
    """
    Клавиатура для выбора раздела архива.
    
    Показывает кнопки:
    - 📦 В архиве (КП со статусом "в архиве")
    - 🏭 В производстве (КП со статусом "в работе")
    - ✅ Выполненные КП (КП со статусом "выполнено")
    - 📊 Актуальный план (сохранённый план производства)
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📦 В архиве", callback_data="archive_section_archived")],
            [InlineKeyboardButton(text="🏭 В производстве", callback_data="archive_section_production")],
            [InlineKeyboardButton(text="✅ Выполненные КП", callback_data="archive_section_completed")],
            [InlineKeyboardButton(text="📊 Актуальный план", callback_data="view_current_plan")],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="archive_back_to_menu")],
        ]
    )


def kp_details_kb(kp_id: int) -> InlineKeyboardMarkup:
    """
    Клавиатура действий для конкретного КП.
    
    Показывает кнопки:
    - Скачать PDF
    - Скачать XLSX
    - Удалить КП
    - Назад к списку
    
    Args:
        kp_id: номер КП для формирования callback_data
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📄 Скачать PDF", callback_data=f"download_pdf_{kp_id}")],
            [InlineKeyboardButton(text="📊 Скачать XLSX", callback_data=f"download_xlsx_{kp_id}")],
            [InlineKeyboardButton(text="🗑️ Удалить КП", callback_data=f"delete_kp_{kp_id}")],
            [InlineKeyboardButton(text="◀️ Назад к списку", callback_data="archive_back_to_sections")],
        ]
    )