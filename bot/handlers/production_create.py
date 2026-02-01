"""Создание нового плана производства и фильтрация КП"""
import logging
import sqlite3
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from ..keyboards import (
    cancel_process_kb, production_filter_kb, main_menu_kb, tracks_choice_kb
)
from ..states import ProductionStates

# Импорт менеджера планов
from .plan_manager import load_plans_metadata

router = Router()


@router.callback_query(F.data == "start_new_planning")
async def start_new_planning(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки 'Начать планирование'.
    
    Спрашивает дату начала плана.
    
    ВАЖНО: Сбрасываем active_plan_id, чтобы при сохранении 
    создавался НОВЫЙ план, а не добавлялись дорожки к существующему!
    """
    # Сбрасываем активный план — при сохранении будет создан НОВЫЙ план
    await state.update_data(active_plan_id=None)
    
    await state.set_state(ProductionStates.waiting_start_date)
    await callback.message.answer(
        "📅 Шаг 1 из 3: С какого числа начинается ваш план?\n\n"
        "Поддерживаемые форматы:\n"
        "• 22 (число текущего месяца)\n"
        "• 22.01.2026 (полная дата)\n"
        "• 2026-01-22 (ISO формат)\n\n"
        "Или оставьте пустым для сегодняшней даты:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.message(ProductionStates.waiting_start_date)
async def receive_start_date(message: Message, state: FSMContext):
    """
    Получаем дату начала плана.
    
    Простыми словами:
    - Парсим дату из разных форматов
    - Сохраняем в state
    - Переходим к вопросу о количестве дорожек
    """
    user_input = message.text.strip()
    
    # Если пустой ввод — используем сегодня
    if not user_input:
        start_date = datetime.now()
        date_description = "сегодня (" + start_date.strftime('%d.%m.%Y') + ")"
    else:
        start_date = None
        date_description = ""
        
        # Формат 1: Полная дата "ДД.ММ.ГГГГ"
        try:
            start_date = datetime.strptime(user_input, '%d.%m.%Y')
            date_description = start_date.strftime('%d.%m.%Y')
        except ValueError:
            pass
        
        # Формат 2: Полная дата "ГГГГ-ММ-ДД"
        if not start_date:
            try:
                start_date = datetime.strptime(user_input, '%Y-%m-%d')
                date_description = start_date.strftime('%d.%m.%Y')
            except ValueError:
                pass
        
        # Формат 3: Только число месяца
        if not start_date:
            try:
                date_number = int(user_input)
                
                if date_number < 1 or date_number > 31:
                    await message.answer(
                        "❌ Число должно быть от 1 до 31.\n"
                        "Попробуйте снова:",
                        reply_markup=cancel_process_kb()
                    )
                    return
                
                now = datetime.now()
                start_date = datetime(now.year, now.month, date_number)
                date_description = start_date.strftime('%d.%m.%Y')
            except ValueError:
                pass
        
        if not start_date:
            await message.answer(
                "❌ Неверный формат даты.\n\n"
                "Поддерживаемые форматы:\n"
                "• 22 (число месяца)\n"
                "• 22.01.2026 (полная дата)\n"
                "• 2026-01-22 (ISO формат)\n\n"
                "Попробуйте снова:",
                reply_markup=cancel_process_kb()
            )
            return
    
    # Сохраняем дату начала
    await state.update_data(
        plan_start_date=start_date.strftime('%Y-%m-%d'),
        plan_start_description=date_description
    )
    
    # Переходим к вопросу о дорожках
    await state.set_state(ProductionStates.waiting_tracks_count)
    await message.answer(
        f"✅ Дата начала плана: {date_description}\n\n"
        "📋 Шаг 2 из 3: Сколько дорожек нужно загрузить в день?\n"
        "(Выберите число или введите вручную)",
        reply_markup=tracks_choice_kb()
    )


@router.message(ProductionStates.waiting_tracks_count)
async def receive_tracks_count(message: Message, state: FSMContext):
    """Получаем количество дорожек (текстовый ввод)"""
    try:
        tracks_count = int(message.text.strip())
        
        if tracks_count <= 0 or tracks_count > 50:
            await message.answer(
                "❌ Количество дорожек должно быть от 1 до 50.\n"
                "Попробуйте снова:"
            )
            await message.answer(
                "Или выберите число на кнопках ниже:",
                reply_markup=tracks_choice_kb()
            )
            return
    except ValueError:
        await message.answer(
            "❌ Неверный формат. Введите целое число (например: 5):"
        )
        await message.answer(
            "Или выберите число на кнопках ниже:",
            reply_markup=tracks_choice_kb()
        )
        return
    
    await state.update_data(tracks_count=tracks_count)
    
    # Переходим к выбору способа фильтрации
    await state.set_state(ProductionStates.waiting_filter_method)
    await message.answer(
        f"✅ Дорожек: {tracks_count}\n\n"
        "📋 Шаг 3 из 3: Как выбрать плиты для производства?\n\n"
        "Выберите способ:",
        reply_markup=production_filter_kb()
    )


@router.callback_query(F.data.startswith("tracks_"), ProductionStates.waiting_tracks_count)
async def process_tracks_choice(callback: CallbackQuery, state: FSMContext):
    """Обработчик нажатия на кнопку с количеством дорожек"""
    tracks_count = int(callback.data.split("_")[1])
    
    await state.update_data(tracks_count=tracks_count)
    
    # Убираем кнопки у текущего сообщения
    await callback.message.edit_reply_markup(reply_markup=None)
    
    # Переходим к выбору способа фильтрации
    await state.set_state(ProductionStates.waiting_filter_method)
    await callback.message.answer(
        f"✅ Дорожек: {tracks_count}\n\n"
        "📋 Шаг 3 из 3: Как выбрать плиты для производства?\n\n"
        "Выберите способ:",
        reply_markup=production_filter_kb()
    )
    await callback.answer()


# === ОБРАБОТЧИКИ ВЫБОРА СПОСОБА ФИЛЬТРАЦИИ ===

@router.callback_query(F.data == "filter_by_date", ProductionStates.waiting_filter_method)
async def filter_by_date(callback: CallbackQuery, state: FSMContext):
    """Выбор по дате - показываем текущую логику"""
    await state.update_data(filter_method='date')
    await state.set_state(ProductionStates.waiting_date_number)
    await callback.message.answer(
        "Шаг 3 из 3: До какой даты брать плиты?\n\n"
        "Поддерживаемые форматы:\n"
        "• 25 (число текущего месяца)\n"
        "• 01.02.2026 (полная дата)\n"
        "• 2026-02-01 (ISO формат)\n\n"
        "Введите дату:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "filter_by_kp", ProductionStates.waiting_filter_method)
async def filter_by_kp(callback: CallbackQuery, state: FSMContext):
    """Выбор по номерам КП"""
    await state.update_data(filter_method='kp')
    await state.set_state(ProductionStates.waiting_kp_numbers)
    await callback.message.answer(
        "Шаг 3 из 3: Введите номера КП\n\n"
        "Поддерживаемые форматы:\n"
        "• 15 (один номер)\n"
        "• 15, 18, 22 (несколько через запятую)\n"
        "• 15-22 (диапазон)\n\n"
        "Введите номера:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "filter_all", ProductionStates.waiting_filter_method)
async def filter_all(callback: CallbackQuery, state: FSMContext):
    """Выбор всех КП в работе - сразу переходим к загрузке"""
    await state.update_data(filter_method='all')
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_execution import load_and_plan_production
    await load_and_plan_production(callback.message, state)
    await callback.answer()


@router.callback_query(F.data == "filter_by_customer", ProductionStates.waiting_filter_method)
async def filter_by_customer(callback: CallbackQuery, state: FSMContext):
    """Выбор по заказчику - показываем список заказчиков"""
    await state.update_data(filter_method='customer')
    
    # Загружаем список уникальных заказчиков из БД
    db_path = PROJECT_ROOT / "plita.db"
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute('''
        SELECT DISTINCT customer_name 
        FROM KP_offers 
        WHERE customer_name IS NOT NULL
        ORDER BY customer_name
    ''')
    customers = [row[0] for row in cur.fetchall()]
    conn.close()
    
    if not customers:
        await callback.message.answer(
            "❌ Нет заказчиков в базе данных.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
        await callback.answer()
        return
    
    # Создаем клавиатуру с заказчиками
    buttons = []
    for customer in customers:
        buttons.append([
            InlineKeyboardButton(
                text=customer,
                callback_data=f"customer_{customer}"
            )
        ])
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад", callback_data="cancel_process")
    ])
    
    await callback.message.answer(
        "Выберите заказчика:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )
    await callback.answer()


@router.callback_query(F.data.startswith("customer_"), ProductionStates.waiting_filter_method)
async def receive_customer_selection(callback: CallbackQuery, state: FSMContext):
    """Обработчик выбора заказчика"""
    customer_name = callback.data.replace("customer_", "")
    
    await state.update_data(customer_name=customer_name)
    
    await callback.message.answer(f"✅ Заказчик: {customer_name}\n\n⏳ Загружаю плиты...")
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_execution import load_and_plan_production
    await load_and_plan_production(callback.message, state)
    await callback.answer()


# === ОБРАБОТЧИКИ ВВОДА НОМЕРОВ КП ===

@router.message(ProductionStates.waiting_kp_numbers)
async def receive_kp_numbers(message: Message, state: FSMContext):
    """Парсим номера КП и запускаем планирование"""
    user_input = message.text.strip()
    
    kp_ids = []
    
    try:
        # Формат 1: Диапазон "15-22"
        if '-' in user_input and user_input.count('-') == 1:
            parts = user_input.split('-')
            start_str = parts[0].strip()
            end_str = parts[1].strip()
            
            # Проверка что это не дата (не содержит точки)
            if '.' not in start_str and '.' not in end_str:
                start_num = int(start_str)
                end_num = int(end_str)
                if start_num > end_num:
                    raise ValueError("Начало диапазона больше конца")
                kp_ids = list(range(start_num, end_num + 1))
            else:
                raise ValueError("Формат не распознан как диапазон номеров КП")
        
        # Формат 2: Список "15, 18, 22"
        elif ',' in user_input:
            kp_ids = [int(x.strip()) for x in user_input.split(',')]
        
        # Формат 3: Одно число "15"
        else:
            kp_ids = [int(user_input)]
        
        if not kp_ids:
            raise ValueError("Не указаны номера КП")
        
    except ValueError as e:
        await message.answer(
            f"❌ Неверный формат: {e}\n\n"
            "Попробуйте снова:\n"
            "• 15 (один номер)\n"
            "• 15, 18, 22 (список)\n"
            "• 15-22 (диапазон)"
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return
    
    # Сохраняем номера КП
    await state.update_data(kp_ids=kp_ids)
    
    await message.answer(f"✅ Выбрано КП: {', '.join(map(str, kp_ids))}\n\n⏳ Загружаю плиты...")
    
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_execution import load_and_plan_production
    await load_and_plan_production(message, state)


@router.message(ProductionStates.waiting_date_number)
async def receive_date_number_and_plan(message: Message, state: FSMContext):
    """Парсим дату и запускаем планирование"""
    user_input = message.text.strip()
    
    # === ПАРСИНГ ДАТЫ: поддерживаем разные форматы ===
    target_date = None
    date_description = ""
    
    # Формат 1: Полная дата "ДД.ММ.ГГГГ"
    try:
        target_date = datetime.strptime(user_input, '%d.%m.%Y')
        date_description = target_date.strftime('%d.%m.%Y')
    except ValueError:
        pass
    
    # Формат 2: Полная дата "ГГГГ-ММ-ДД"
    if not target_date:
        try:
            target_date = datetime.strptime(user_input, '%Y-%m-%d')
            date_description = target_date.strftime('%d.%m.%Y')
        except ValueError:
            pass
    
    # Формат 3: Только число месяца
    if not target_date:
        try:
            date_number = int(user_input)
            
            if date_number < 1 or date_number > 31:
                await message.answer(
                    "❌ Число должно быть от 1 до 31.\n"
                    "Попробуйте снова:"
                )
                await message.answer(
                    "Или нажмите кнопку ниже для отмены:",
                    reply_markup=cancel_process_kb()
                )
                return
            
            now = datetime.now()
            target_date = datetime(now.year, now.month, date_number)
            date_description = f"{date_number} {target_date.strftime('%B %Y')}"
        except ValueError:
            pass
    
    if not target_date:
        await message.answer(
            "❌ Неверный формат даты.\n\n"
            "Поддерживаемые форматы:\n"
            "• 25 (число месяца)\n"
            "• 01.02.2026 (полная дата)\n"
            "• 2026-02-01 (ISO формат)\n\n"
            "Попробуйте снова:"
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return
    
    # Сохраняем дату и описание
    await state.update_data(
        target_date=target_date.isoformat(),
        date_description=date_description
    )
    
    data = await state.get_data()
    tracks_count = data.get('tracks_count', 1)
    
    await message.answer(
        f"✅ Параметры планирования:\n"
        f"• Дорожек в день: {tracks_count}\n"
        f"• Плиты со сроком до: {date_description}\n\n"
        f"⏳ Загружаю плиты из базы данных..."
    )
    
    # Импортируем здесь, чтобы избежать циклических импортов
    from .production_execution import load_and_plan_production
    await load_and_plan_production(message, state)
