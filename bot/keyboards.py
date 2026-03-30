"""Клавиатуры бота"""
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton


def main_menu_kb() -> ReplyKeyboardMarkup:
    """Главное меню бота"""
    return ReplyKeyboardMarkup(
        keyboard=[
            # Первая строка - 2 кнопки рядом
            [
                KeyboardButton(text="📝 Создать КП"),
                KeyboardButton(text="Планирование производства")
            ],
            # Вторая строка - 2 кнопки рядом
            [
                KeyboardButton(text="📁 Архив"),
                KeyboardButton(text="Информация о ПБ в работе")
            ],
            # Третья строка - 2 кнопки рядом
            [
                KeyboardButton(text="📖 Как работает бот"),
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
            [InlineKeyboardButton(text="📊 КП в производстве", callback_data="kp_in_production")],
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


def save_to_db_with_files_kb(
    has_pdf: bool = False,
    has_xlsx: bool = False,
    has_breakdown: bool = False,
    has_schema: bool = False,
    has_schema_breakdown: bool = False,
) -> InlineKeyboardMarkup:
    """
    Объединённая клавиатура: кнопки скачивания файлов + сохранение в БД.
    Кнопки файлов показываются только для существующих документов.
    """
    buttons = []
    # Ряд с кнопками файлов (по 2-3 в ряд)
    file_row = []
    if has_pdf:
        file_row.append(InlineKeyboardButton(text="📄 PDF", callback_data="kp_file_pdf"))
    if has_xlsx:
        file_row.append(InlineKeyboardButton(text="📊 XLSX", callback_data="kp_file_xlsx"))
    if has_breakdown:
        file_row.append(InlineKeyboardButton(text="📋 Разбивка", callback_data="kp_file_breakdown"))
    if file_row:
        buttons.append(file_row)
    second_file_row = []
    if has_schema:
        second_file_row.append(InlineKeyboardButton(text="📐 Схема", callback_data="kp_file_schema"))
    if has_schema_breakdown:
        second_file_row.append(InlineKeyboardButton(text="📊 Разб.схемы", callback_data="kp_file_schema_breakdown"))
    if second_file_row:
        buttons.append(second_file_row)
    # Кнопки сохранения
    buttons.append([InlineKeyboardButton(text="💾 Сохранить в БД", callback_data="save_kp_to_db")])
    buttons.append([InlineKeyboardButton(text="📦 В архив", callback_data="save_kp_to_archive")])
    buttons.append([InlineKeyboardButton(text="❌ Не сохранять", callback_data="skip_save_kp")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def production_menu_kb() -> InlineKeyboardMarkup:
    """
    Начальное меню планирования производства.
    
    Показывает кнопки:
    - Календарный план — просмотр активного плана с датами
    - Начать планирование — создание нового плана
    - Планы — просмотр всех сохранённых планов
    - Производственный календарь — управление рабочими и нерабочими днями
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📅 Календарный план", callback_data="view_calendar_plan")],
            [InlineKeyboardButton(text="🚀 Начать планирование", callback_data="start_new_planning")],
            [InlineKeyboardButton(text="📋 Планы", callback_data="view_all_plans")],
            [InlineKeyboardButton(text="🗓️ Производственный календарь", callback_data="manage_work_calendar")],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")],
        ]
    )


def production_days_kb(total_days: int) -> InlineKeyboardMarkup:
    """
    Создает клавиатуру с кнопками для выбора дня производства
    (старая версия, оставлена для совместимости)
    
    Args:
        total_days: общее количество дней работы
        
    Returns:
        InlineKeyboardMarkup с кнопками "День 1", "День 2" и т.д.
    """
    buttons = []
    
    # Кнопка диаграммы текущего плана (этап сохранения — только «этого плана»)
    buttons.append([
        InlineKeyboardButton(
            text="📈 Диаграмма этого плана",
            callback_data="export_gantt_current"
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
    
    # Кнопка "Сохранить план"
    buttons.append([
        InlineKeyboardButton(
            text="💾 Сохранить план",
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


def calendar_days_kb(
    total_days: int, 
    start_date: str, 
    completed_days: list = None,
    days_info: dict = None,
    show_save_button: bool = True
) -> InlineKeyboardMarkup:
    """
    Создает клавиатуру с ДАТАМИ для выбора дня производства.
    
    Простыми словами:
    - Вместо "День 1, День 2" показывает "22.01, 23.01, 24.01..."
    - Выполненные дни отмечаются галочкой ✅
    - Невыполненные — значком 📅
    - Показывает ГЛОБАЛЬНУЮ загруженность: "22.01 3/5" (3 дорожки занято во ВСЕХ планах / 5 максимум)
    
    Args:
        total_days: общее количество дней работы
        start_date: дата начала плана в формате "YYYY-MM-DD" или "DD.MM.YYYY"
        completed_days: список номеров выполненных дней [1, 2, 3]
        days_info: информация о днях с ГЛОБАЛЬНОЙ загруженностью {
            "2026-01-22": {"occupied": 3, "max": 5, "completed": False, "day_number": 1},
            ...
        }
        show_save_button: показывать кнопку "Сохранить план" (по умолчанию True).
            При True показывается только «Диаграмма этого плана», при False — только «Диаграмма Ганта» (все планы).
        
    Returns:
        InlineKeyboardMarkup с кнопками-датами
    """
    from datetime import datetime
    from core.work_calendar import nth_working_day
    
    # Константа максимума дорожек (импортируем здесь чтобы избежать циклических импортов)
    MAX_TRACKS = 5
    
    if completed_days is None:
        completed_days = []
    
    if days_info is None:
        days_info = {}
    
    # Парсим дату начала
    parsed_start = None
    if start_date:
        # Пробуем разные форматы
        for fmt in ['%Y-%m-%d', '%d.%m.%Y', '%Y-%m-%dT%H:%M:%S']:
            try:
                parsed_start = datetime.strptime(start_date.split('T')[0] if 'T' in start_date else start_date, fmt.split('T')[0])
                break
            except ValueError:
                continue
    
    # Если не удалось распарсить — используем сегодня
    if not parsed_start:
        parsed_start = datetime.now()
    
    buttons = []
    
    # Одна кнопка диаграммы: на этапе сохранения — «этого плана», в календаре — «Ганта» (все планы)
    if show_save_button:
        buttons.append([
            InlineKeyboardButton(
                text="📈 Диаграмма этого плана",
                callback_data="export_gantt_current"
            )
        ])
    else:
        buttons.append([
            InlineKeyboardButton(
                text="📊 Диаграмма Ганта",
                callback_data="export_gantt"
            )
        ])
    
    # Получаем текущую дату для фильтрации прошедших дней
    today = datetime.now().date()
    
    # Создаем кнопки по 3 в ряд для компактности
    row = []
    for day in range(1, total_days + 1):
        # Вычисляем дату этого дня
        day_date = datetime.combine(
            nth_working_day(parsed_start.date(), day),
            datetime.min.time(),
        )
        date_str = day_date.strftime("%d.%m")  # Формат: "22.01"
        date_key = day_date.strftime("%Y-%m-%d")  # Формат: "2026-01-22"
        
        # Получаем информацию о дне из days_info
        day_data = days_info.get(date_key, {})
        
        # ГЛОБАЛЬНАЯ загруженность: occupied = занято во всех планах
        occupied_tracks = day_data.get('occupied', 0)
        max_tracks = day_data.get('max', MAX_TRACKS)
        is_completed = day_data.get('completed', False)
        
        # Скрываем выполненные прошедшие дни
        if (day in completed_days or is_completed) and day_date.date() < today:
            continue
        
        # Определяем эмодзи: галочка если выполнен, календарь если нет
        if day in completed_days or is_completed:
            emoji = "✅"
        else:
            emoji = "📅"
        
        # Формируем текст кнопки с ГЛОБАЛЬНОЙ загруженностью
        if occupied_tracks > 0:
            # Показываем загруженность: "22.01 3/5" (занято/максимум)
            button_text = f"{emoji} {date_str} {occupied_tracks}/{max_tracks}"
        else:
            # Обычный формат без загруженности
            button_text = f"{emoji} {date_str}"
        
        row.append(InlineKeyboardButton(
            text=button_text,
            callback_data=f"production_day_{day}"
        ))
        
        # Каждые 3 кнопки - новый ряд
        if len(row) == 3:
            buttons.append(row)
            row = []
    
    # Добавляем оставшиеся кнопки
    if row:
        buttons.append(row)
    
    # Кнопка "Сохранить план" (только если нужна)
    if show_save_button:
        buttons.append([
            InlineKeyboardButton(
                text="💾 Сохранить план",
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


def production_day_actions_kb(
    day_number: int, 
    total_days: int, 
    start_date: str = None, 
    completed_days: list = None,
    show_save_button: bool = True
) -> InlineKeyboardMarkup:
    """
    Клавиатура действий после просмотра дня производства.
    Показывает кнопку "Выполнено" и кнопки для перехода на другие дни.
    
    Args:
        day_number: номер текущего дня
        total_days: общее количество дней
        start_date: дата начала плана (опционально)
        completed_days: список выполненных дней (опционально)
        show_save_button: показывать кнопку "Сохранить план" (по умолчанию True)
        
    Returns:
        InlineKeyboardMarkup с кнопками действий
    """
    from datetime import datetime
    from core.work_calendar import nth_working_day
    
    if completed_days is None:
        completed_days = []
    
    # Парсим дату начала
    parsed_start = None
    if start_date:
        for fmt in ['%Y-%m-%d', '%d.%m.%Y']:
            try:
                parsed_start = datetime.strptime(start_date.split('T')[0] if 'T' in start_date else start_date, fmt)
                break
            except ValueError:
                continue
    
    if not parsed_start:
        parsed_start = datetime.now()
    
    # Получаем текущую дату для фильтрации прошедших дней
    today = datetime.now().date()
    
    buttons = []
    
    # Кнопка "День выполнен" для текущего дня
    buttons.append([
        InlineKeyboardButton(
            text="✅ День выполнен",
            callback_data=f"complete_day_{day_number}"
        )
    ])
    
    # Кнопка "Сохранить план" (только если нужна)
    if show_save_button:
        buttons.append([
            InlineKeyboardButton(
                text="💾 Сохранить план",
                callback_data="save_current_plan"
            )
        ])
    
    # Кнопки навигации по дням (по 3 в ряд) с датами
    row = []
    for day in range(1, total_days + 1):
        # Вычисляем дату
        day_date = datetime.combine(
            nth_working_day(parsed_start.date(), day),
            datetime.min.time(),
        )
        date_str = day_date.strftime("%d.%m")
        
        # Определяем эмодзи
        if day in completed_days:
            emoji = "✅"
        else:
            emoji = "📅"
        
        # Скрываем выполненные прошедшие дни
        if day in completed_days and day_date.date() < today:
            continue
        
        row.append(InlineKeyboardButton(
            text=f"{emoji} {date_str}",
            callback_data=f"production_day_{day}"
        ))
        if len(row) == 3:
            buttons.append(row)
            row = []
    if row:
        buttons.append(row)
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def plates_completion_kb(plates_by_track: list, rejected_quantities: dict, active_plate_id: tuple = None) -> InlineKeyboardMarkup:
    """
    Клавиатура для отметки бракованных плит при завершении дня.
    
    Простыми словами:
    - Показывает список плит по дорожкам с кнопками
    - Каждая дорожка начинается с заголовка
    - Клик на плиту открывает счетчик брака под ней
    - Счетчик позволяет выбрать количество бракованных плит кнопками +/-
    - Плиты из остатков показываются отдельной группой с пометкой
    - Внизу кнопки "Подтвердить" и "Отмена"
    
    Args:
        plates_by_track: список дорожек, каждая содержит track_number и список plates
        rejected_quantities: словарь {(track_idx, plate_idx): количество_брака}
        active_plate_id: кортеж (track_idx, plate_idx) активной плиты или None
        
    Returns:
        InlineKeyboardMarkup с кнопками плит
    """
    buttons = []
    
    for track_idx, track_data in enumerate(plates_by_track):
        track_number = track_data.get('track_number', track_idx + 1)
        plates = track_data.get('plates', [])
        is_from_rests = track_data.get('is_from_rests', False)
        
        if not plates:
            continue
        
        # Добавляем заголовок дорожки (кнопка без действия)
        # Для плит из остатков показываем специальный заголовок
        if is_from_rests or track_number == 0:
            header_text = "━━━━ Из остатков ━━━━"
        else:
            header_text = f"━━━━ Дорожка {track_number} ━━━━"
        
        buttons.append([
            InlineKeyboardButton(
                text=header_text,
                callback_data=f"track_header_{track_idx}"
            )
        ])
        
        # Добавляем плиты этой дорожки
        for plate_idx, plate in enumerate(plates):
            plate_id = (track_idx, plate_idx)
            
            # Получаем количество брака для этой плиты
            reject_qty = rejected_quantities.get(plate_id, 0)
            total_qty = plate.get('qty', 1)
            
            # Если есть брак — показываем ❌, иначе ✅
            emoji = "❌" if reject_qty > 0 else "✅"
            
            plate_name = plate.get('plate_name', f"Плита {plate_idx+1}")
            kp_date = plate.get('kp_date', '')
            kp_id = plate.get('kp_id', '')
            from_rest = plate.get('from_rest', False)
            match_type = plate.get('match_type', '')
            is_secondary = plate.get('is_secondary', False)  # Флаг вторичного реза
            
            # Форматируем дату коротко: "02.02.2026" -> "02.02"
            date_short = kp_date[:5] if kp_date and kp_date != 'неизвестно' else ''
            
            # Формируем информацию о КП
            if kp_id and date_short:
                kp_info = f"({date_short})"
            elif kp_id:
                kp_info = f"(КП{kp_id})"
            else:
                kp_info = ""
            
            # Для плит из остатков добавляем пометку о типе
            if from_rest:
                if match_type == 'exact':
                    rest_mark = "[=]"  # Точное совпадение
                else:
                    rest_mark = "[рез]"  # Нужен рез
                # Показываем полное название для плит из остатков
                max_name_len = 50
            else:
                rest_mark = ""
                max_name_len = 18 if kp_info else 25
            
            # Добавляем метку для вторичных резов
            if is_secondary:
                plate_name = f"[2] {plate_name}"  # [2] = вторичный рез (из остатка основной плиты)
                max_name_len += 4  # Учитываем длину метки
            
            # Обрезаем длинные названия
            if len(plate_name) > max_name_len:
                plate_name = plate_name[:max_name_len-2] + ".."
            
            # Формируем текст кнопки
            if from_rest:
                button_text = f"{emoji} {rest_mark} {plate_name} {kp_info} ×{total_qty}"
            else:
                button_text = f"{emoji} {plate_name} {kp_info} × {total_qty}"
            
            # Основная кнопка плиты
            buttons.append([
                InlineKeyboardButton(
                    text=button_text,
                    callback_data=f"plate_open_t{track_idx}_p{plate_idx}"
                )
            ])
            
            # Если это активная плита — добавляем строку управления браком
            if active_plate_id == plate_id:
                control_buttons = [
                    InlineKeyboardButton(
                        text="−",
                        callback_data=f"reject_minus_t{track_idx}_p{plate_idx}"
                    ),
                    InlineKeyboardButton(
                        text=f"Брак: {reject_qty}/{total_qty}",
                        callback_data=f"reject_info_t{track_idx}_p{plate_idx}"
                    ),
                    InlineKeyboardButton(
                        text="+",
                        callback_data=f"reject_plus_t{track_idx}_p{plate_idx}"
                    ),
                    InlineKeyboardButton(
                        text="🔄 Сбросить",
                        callback_data=f"reject_reset_t{track_idx}_p{plate_idx}"
                    )
                ]
                buttons.append(control_buttons)
    
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
    - Просмотр остатков (экспорт в Excel)
    - Статистика БД
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🗑️ Очистить все данные", callback_data="db_clear_all")],
            [InlineKeyboardButton(text="📋 Просмотр остатков", callback_data="db_view_rests")],
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


def confirm_plates_list_kb() -> InlineKeyboardMarkup:
    """Кнопки подтверждения/замены списка плит и донажатия позиций"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Подтвердить", callback_data="confirm_plates_list"),
                InlineKeyboardButton(text="🔄 Заменить", callback_data="replace_plates_list"),
            ],
            [
                InlineKeyboardButton(text="➕ Продолжить КП", callback_data="continue_kp_plates"),
            ],
        ]
    )


def wide_plates_actions_kb() -> InlineKeyboardMarkup:
    """Кнопки для шага обработки плит шире 12 дм."""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="⏭️ Пропустить широкие плиты", callback_data="skip_wide_plates")],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")],
        ]
    )


def tracks_choice_kb() -> InlineKeyboardMarkup:
    """Клавиатура выбора количества дорожек (1-5)"""
    buttons = []
    # Создаем кнопки от 1 до 5 в один ряд
    row = []
    for i in range(1, 6):
        row.append(InlineKeyboardButton(text=str(i), callback_data=f"tracks_{i}"))
    buttons.append(row)
    # Шаг назад (на шаг 1) + Назад в меню
    buttons.append([InlineKeyboardButton(text="◀️ Шаг назад", callback_data="plan_step_back_to_1")])
    buttons.append([InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def production_filter_kb() -> InlineKeyboardMarkup:
    """Клавиатура выбора способа фильтрации плит для производства"""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📅 По дате", callback_data="filter_by_date")],
            [InlineKeyboardButton(text="📋 По КП", callback_data="filter_by_kp_buttons")],
            [InlineKeyboardButton(text="📦 Все КП в работе", callback_data="filter_all")],
            [InlineKeyboardButton(text="👤 По заказчику", callback_data="filter_by_customer")],
            [InlineKeyboardButton(text="◀️ Шаг назад", callback_data="plan_step_back_to_2")],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="cancel_process")],
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
            [InlineKeyboardButton(text="🔍 Найти по номеру КП", callback_data="archive_find_by_number")],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="archive_back_to_menu")],
        ]
    )


def kp_details_kb(kp_id: int, total_amount: float = 0, status: str | None = None) -> InlineKeyboardMarkup:
    """
    Клавиатура действий для конкретного КП.

    Показывает кнопки:
    - PDF / XLSX с суммой в подписи (если total_amount > 0)
    - Изменить скидку
    - Удалить КП
    - В производство (только если status == "в архиве")
    - Назад к списку

    Args:
        kp_id: номер КП для формирования callback_data
        total_amount: итоговая сумма с НДС для отображения в кнопках PDF/XLSX
        status: статус КП; при "в архиве" добавляется кнопка "В производство"
    """
    amount_str = f"{total_amount:,.0f} ₽".replace(",", " ") if total_amount else ""
    pdf_text = f"📄 PDF · {amount_str}" if amount_str else "📄 PDF"
    xlsx_text = f"📊 XLSX · {amount_str}" if amount_str else "📊 XLSX"
    rows = [
        [InlineKeyboardButton(text=pdf_text, callback_data=f"download_pdf_{kp_id}")],
        [InlineKeyboardButton(text=xlsx_text, callback_data=f"download_xlsx_{kp_id}")],
        [InlineKeyboardButton(text="✏️ Изменить скидку", callback_data=f"change_discount_{kp_id}")],
        [InlineKeyboardButton(text="🗑️ Удалить КП", callback_data=f"delete_kp_{kp_id}")],
    ]
    if status == "в архиве":
        rows.append([InlineKeyboardButton(text="🏭 В производство", callback_data=f"move_kp_to_production_{kp_id}")])
    rows.append([InlineKeyboardButton(text="◀️ Назад к списку", callback_data="archive_back_to_sections")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def kp_production_details_kb(kp_id: int) -> InlineKeyboardMarkup:
    """
    Клавиатура для детального просмотра КП в производстве.
    
    Показывает кнопки:
    - Скачать PDF
    - Скачать XLSX
    - Изменить дату (НОВОЕ!)
    - Назад к списку
    
    Args:
        kp_id: номер КП для формирования callback_data
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📄 Скачать PDF", callback_data=f"download_pdf_{kp_id}")],
            [InlineKeyboardButton(text="📊 Скачать XLSX", callback_data=f"download_xlsx_{kp_id}")],
            [InlineKeyboardButton(text="📅 Изменить дату", callback_data=f"change_date_{kp_id}")],
            [InlineKeyboardButton(text="◀️ Назад к списку", callback_data="kp_in_production")],
        ]
    )


def instructions_choice_kb() -> InlineKeyboardMarkup:
    """
    Клавиатура для выбора роли (Менеджер или Производство).
    
    Показывает кнопки:
    - 👨‍💼 Для Менеджера - инструкция по работе с КП и архивом
    - 🏭 Для Производства - инструкция по планированию и отчётам
    - ◀️ Назад в меню - возврат в главное меню
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="👨‍💼 Для Менеджера", callback_data="instructions_manager")],
            [InlineKeyboardButton(text="🏭 Для Производства", callback_data="instructions_production")],
            [InlineKeyboardButton(text="◀️ Назад в меню", callback_data="instructions_back_to_menu")],
        ]
    )


def instructions_back_kb() -> InlineKeyboardMarkup:
    """
    Клавиатура для возврата после просмотра инструкции.
    
    Показывает кнопки:
    - 🔙 К выбору роли - вернуться к выбору Менеджер/Производство
    - 🏠 В главное меню - вернуться в начало
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🔙 К выбору роли", callback_data="instructions_back_to_choice")],
            [InlineKeyboardButton(text="🏠 В главное меню", callback_data="instructions_back_to_menu")],
        ]
    )


def plans_list_kb(plans_list: list, active_plan_id: str = None) -> InlineKeyboardMarkup:
    """
    Клавиатура со списком всех сохранённых планов.
    
    Простыми словами:
    - Показывает список планов с их названиями
    - Активный план помечается звёздочкой ⭐
    - Внизу кнопки "Создать новый план" и "Назад"
    
    Args:
        plans_list: список планов из plans_metadata.json
        active_plan_id: ID активного плана (помечается звёздочкой)
        
    Returns:
        InlineKeyboardMarkup с кнопками планов
    """
    buttons = []
    
    if not plans_list:
        # Если планов нет — показываем сообщение
        buttons.append([
            InlineKeyboardButton(
                text="📭 Нет сохранённых планов",
                callback_data="no_plans_info"
            )
        ])
    else:
        # Показываем список планов
        for plan in plans_list:
            plan_id = plan.get('id', '')
            plan_name = plan.get('name', f'План {plan_id}')
            total_days = plan.get('total_days', 0)
            total_tracks = plan.get('total_tracks', 0)
            
            # Помечаем активный план звёздочкой
            if plan_id == active_plan_id:
                emoji = "⭐"
            else:
                emoji = "📋"
            
            # Формируем текст кнопки
            button_text = f"{emoji} {plan_name} ({total_days}д, {total_tracks} дор.)"
            
            # Обрезаем если слишком длинный
            if len(button_text) > 50:
                button_text = button_text[:47] + "..."
            
            buttons.append([
                InlineKeyboardButton(
                    text=button_text,
                    callback_data=f"select_plan_{plan_id}"
                )
            ])
    
    # Кнопка создания нового плана
    buttons.append([
        InlineKeyboardButton(
            text="➕ Создать новый план",
            callback_data="create_new_plan"
        )
    ])
    
    # Кнопка "Назад"
    buttons.append([
        InlineKeyboardButton(
            text="◀️ Назад",
            callback_data="back_to_production_menu"
        )
    ])
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def plan_actions_kb(plan_id: str, is_active: bool = False) -> InlineKeyboardMarkup:
    """
    Клавиатура действий с выбранным планом.
    
    Показывает кнопки:
    - Открыть календарь плана
    - Сделать активным (если не активный)
    - Удалить план
    - Назад к списку
    
    Args:
        plan_id: ID плана
        is_active: является ли план активным
    """
    buttons = []
    
    # Открыть календарь
    buttons.append([
        InlineKeyboardButton(
            text="📅 Открыть календарь",
            callback_data=f"open_plan_calendar_{plan_id}"
        )
    ])
    
    # Сделать активным (если не активный)
    if not is_active:
        buttons.append([
            InlineKeyboardButton(
                text="⭐ Сделать активным",
                callback_data=f"activate_plan_{plan_id}"
            )
        ])
    
    # Удалить план
    buttons.append([
        InlineKeyboardButton(
            text="🗑️ Удалить план",
            callback_data=f"delete_plan_{plan_id}"
        )
    ])
    
    # Назад к списку
    buttons.append([
        InlineKeyboardButton(
            text="◀️ К списку планов",
            callback_data="view_all_plans"
        )
    ])
    
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def confirm_delete_plan_kb(plan_id: str) -> InlineKeyboardMarkup:
    """
    Клавиатура подтверждения удаления плана.
    
    Args:
        plan_id: ID плана для удаления
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(
                text="⚠️ Да, удалить план",
                callback_data=f"confirm_delete_plan_{plan_id}"
            )],
            [InlineKeyboardButton(
                text="❌ Отмена",
                callback_data="view_all_plans"
            )],
        ]
    )


def day_documents_menu_kb(day_number: int, track_numbers: str) -> InlineKeyboardMarkup:
    """
    Меню выбора типа документа для дня производства.
    
    Простыми словами:
    - После просмотра состава дня показывает кнопки
    - Каждая кнопка генерирует определённый тип документа
    - Пользователь выбирает только нужные документы
    
    Args:
        day_number: номер дня в плане (например, 3)
        track_numbers: строка с номерами дорожек (например, "7-9" или "7")
        
    Returns:
        InlineKeyboardMarkup с кнопками выбора документов
    """
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(
                text=f"📐 Схема дорожек {track_numbers}",
                callback_data=f"generate_schema_{day_number}"
            )],
            [InlineKeyboardButton(
                text=f"📊 Детальная разбивка",
                callback_data=f"generate_breakdown_{day_number}"
            )],
            [InlineKeyboardButton(
                text=f"📋 Файлы формовки",
                callback_data=f"generate_formovka_{day_number}"
            )],
            [InlineKeyboardButton(
                text="✅ День выполнен",
                callback_data=f"complete_day_{day_number}"
            )],
            [InlineKeyboardButton(
                text="◀️ Назад к календарю",
                callback_data="back_to_calendar"
            )],
        ]
    )