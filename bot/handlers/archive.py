"""Обработчики для архива коммерческих предложений"""
import asyncio
import json
import logging
import os
import re
import tempfile
from pathlib import Path
from datetime import datetime, timedelta
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, FSInputFile, BufferedInputFile, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.context import FSMContext

from core import kp_db
from core.commercial_offer import generate_commercial_offer_pdf
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.gantt_excel import create_gantt_excel
from core.kp_offer_utils import append_transport_to_order_data, format_offer_quantity
from ..bot_config import OUTPUTS_DIR_STR
from ..keyboards import main_menu_kb, archive_sections_kb, kp_details_kb
from ..states import ArchiveStates
from .plan_manager import get_all_plans_gantt_data

logger = logging.getLogger(__name__)

# Определяем пути к директориям проекта
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent

router = Router()


def _order_data_from_kp_info(kp_info: dict) -> list:
    """Собирает order_data для генераторов из kp_info['plates'] (unit_price из колонки или из discounted_price и скидки)."""
    plates = kp_info.get("plates") or []
    discount = kp_info.get("discount_percent") or 0
    factor = 1.0 - (discount / 100.0)
    if factor <= 0:
        factor = 1.0
    order_data = []
    for p in plates:
        unit_price = p.get("unit_price")
        if unit_price is None or (isinstance(unit_price, (int, float)) and unit_price <= 0):
            discounted_price = p.get("discounted_price") or 0
            unit_price = discounted_price / factor
        qty = p.get("qty") or 0
        total_weight = p.get("total_weight")
        unit_weight = p.get("unit_weight")
        weight = total_weight if total_weight is not None and total_weight > 0 else (unit_weight or 0) * qty
        order_data.append({
            "name": p.get("plate_name") or "",
            "length_m": p.get("length_m") or 0,
            "width_m": p.get("width_m") or 0,
            "qty": qty,
            "load_class": p.get("load_class") or 800,
            "unit_price": float(unit_price),
            "weight": weight or 0,
        })
    return append_transport_to_order_data(
        order_data,
        kp_info.get("transport_hours"),
        kp_info.get("transport_price_per_hour"),
    )


def _parse_positive_number(raw_value: str) -> float:
    value = float((raw_value or "").strip().replace(",", "."))
    if value <= 0:
        raise ValueError
    return value


@router.message(F.text == "📁 Архив")
async def btn_archive(message: Message):
    """
    Обработчик кнопки 'Архив' в главном меню.
    
    Простыми словами:
    - Показывает меню выбора раздела архива
    - Два раздела: "📦 В архиве" и "🏭 В производстве"
    """
    await message.answer(
        "📁 Архив коммерческих предложений\n\n"
        "Выберите раздел для просмотра:",
        reply_markup=archive_sections_kb()
    )


@router.callback_query(F.data == "archive_section_archived")
async def show_archived_kp(callback: CallbackQuery):
    """
    Показать КП в архиве (статус "в архиве").
    
    Простыми словами:
    - Получает все КП со статусом "в архиве" из БД
    - Формирует список с кнопками для каждого КП
    - Сортирует по номеру КП (от меньшего к большему)
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    # Получаем список КП со статусом "в архиве"
    all_kp = kp_db.get_all_kp_list()
    archived_kp = all_kp.get('archived', [])
    
    if not archived_kp:
        await callback.message.edit_text(
            "📦 В архиве пока нет КП\n\n"
            "Чтобы добавить КП в архив, создайте его через '📝 Создать КП' "
            "и нажмите '📦 В архив' после генерации документов.",
            reply_markup=archive_sections_kb()
        )
        return
    
    # Формируем текст с информацией о КП
    text = f"📦 КП в архиве ({len(archived_kp)} шт.)\n\n"
    
    # Создаём inline кнопки для каждого КП
    buttons = []
    
    for kp in archived_kp:
        kp_id = kp['kp_id']
        customer = kp.get('customer_name', 'Без имени')
        total = kp.get('total_amount', 0)
        discount_percent = kp.get('discount_percent', 0)
        
        # Обрезаем длинные имена клиентов
        customer_short = customer[:20] + '...' if len(customer) > 20 else customer
        
        # Формируем текст кнопки с процентом скидки
        buttons.append([
            InlineKeyboardButton(
                text=f"КП №{kp_id} | {customer_short} | {discount_percent:.0f}% | {total:,.0f}₽",
                callback_data=f"view_kp_{kp_id}"
            )
        ])
    
    # Кнопка "Назад"
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")
    ])
    
    await callback.message.edit_text(
        text,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )


@router.callback_query(F.data == "archive_find_by_number")
async def archive_find_by_number(callback: CallbackQuery, state: FSMContext):
    """Запрос номера КП для поиска в архиве."""
    try:
        await callback.answer()
    except Exception:
        pass
    await state.set_state(ArchiveStates.waiting_kp_number)
    back_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")]
    ])
    await callback.message.edit_text(
        "Введите номер КП:\n\nИли нажмите «Назад».",
        reply_markup=back_kb
    )


@router.message(ArchiveStates.waiting_kp_number)
async def receive_kp_number_for_search(message: Message, state: FSMContext):
    """Обработка введённого номера КП при поиске в архиве."""
    raw = (message.text or "").strip()
    try:
        kp_id = int(raw)
    except ValueError:
        await message.answer("Введите число — номер КП.")
        return
    kp_info = kp_db.get_kp_by_id(kp_id)
    back_kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")]
    ])
    if not kp_info:
        await state.clear()
        await message.answer(
            f"КП №{kp_id} не найдено.",
            reply_markup=back_kb
        )
        return
    customer = kp_info.get('customer_name', 'Без имени')
    total = kp_info.get('total_amount', 0)
    discount_percent = kp_info.get('discount_percent', 0)
    customer_short = customer[:20] + '...' if len(customer) > 20 else customer
    buttons = [
        [InlineKeyboardButton(
            text=f"КП №{kp_id} | {customer_short} | {discount_percent:.0f}% | {total:,.0f}₽",
            callback_data=f"view_kp_{kp_id}"
        )],
        [InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")],
    ]
    await state.clear()
    await message.answer(
        "Найденное КП:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )


@router.callback_query(F.data == "archive_section_production")
async def show_production_kp(callback: CallbackQuery):
    """
    Показать КП в производстве (статус "в работе").
    
    Простыми словами:
    - Получает все КП со статусом "в работе" из БД
    - Формирует список с кнопками для каждого КП
    - Сортирует по номеру КП (от меньшего к большему)
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    # Получаем список КП со статусом "в работе"
    all_kp = kp_db.get_all_kp_list()
    production_kp = all_kp.get('in_production', [])
    
    if not production_kp:
        await callback.message.edit_text(
            "🏭 В производстве пока нет КП\n\n"
            "Чтобы отправить КП в производство, создайте его через '📝 Создать КП' "
            "и нажмите '💾 Сохранить в БД' после генерации документов.",
            reply_markup=archive_sections_kb()
        )
        return
    
    # Формируем текст с информацией о КП
    text = f"🏭 КП в производстве ({len(production_kp)} шт.)\n\n"
    
    # Создаём inline кнопки для каждого КП
    buttons = []
    db_path = PROJECT_ROOT / "plita.db"
    
    for kp in production_kp:
        kp_id = kp['kp_id']
        customer = kp.get('customer_name', 'Без имени')
        total = kp.get('total_amount', 0)
        execution_terms = kp.get('execution_terms', '')
        
        # Получаем процент выполнения
        completion_info = kp_db.get_kp_completion_percentage(kp_id, str(db_path))
        percentage = completion_info['percentage']
        
        # Обрезаем длинные имена клиентов
        customer_short = customer[:20] + '...' if len(customer) > 20 else customer
        
        # Формируем текст кнопки с процентом
        button_text = f"КП №{kp_id} | {customer_short} | {percentage:.0f}% | {total:,.0f}₽"
        if execution_terms:
            button_text += f" | ⏰{execution_terms}"
        
        buttons.append([
            InlineKeyboardButton(
                text=button_text,
                callback_data=f"view_kp_{kp_id}"
            )
        ])
    
    # Кнопка "Назад"
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")
    ])
    
    await callback.message.edit_text(
        text,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )


@router.callback_query(F.data == "archive_section_completed")
async def show_completed_kp(callback: CallbackQuery):
    """
    Показать выполненные КП (статус "выполнено").
    
    Простыми словами:
    - Получает все КП со статусом "выполнено" из БД
    - Формирует список с кнопками для каждого КП
    - Сортирует по номеру КП (от меньшего к большему)
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    # Получаем список КП со статусом "выполнено"
    all_kp = kp_db.get_all_kp_list()
    completed_kp = all_kp.get('completed', [])
    
    if not completed_kp:
        await callback.message.edit_text(
            "✅ Выполненных КП пока нет\n\n"
            "КП автоматически получают статус 'выполнено', "
            "когда все плиты из заказа отмечены как выполненные в производстве.",
            reply_markup=archive_sections_kb()
        )
        return
    
    # Формируем текст с информацией о КП
    text = f"✅ Выполненные КП ({len(completed_kp)} шт.)\n\n"
    
    # Создаём inline кнопки для каждого КП
    buttons = []
    db_path = PROJECT_ROOT / "plita.db"
    
    for kp in completed_kp:
        kp_id = kp['kp_id']
        customer = kp.get('customer_name', 'Без имени')
        total = kp.get('total_amount', 0)
        creation_date = kp.get('creation_date', '')
        
        # Получаем процент выполнения (должно быть 100%)
        completion_info = kp_db.get_kp_completion_percentage(kp_id, str(db_path))
        percentage = completion_info['percentage']
        
        # Обрезаем длинные имена клиентов
        customer_short = customer[:20] + '...' if len(customer) > 20 else customer
        
        # Формируем текст кнопки
        button_text = f"КП №{kp_id} | {customer_short} | {percentage:.0f}% | {total:,.0f}₽"
        if creation_date:
            button_text += f" | 📅{creation_date}"
        
        buttons.append([
            InlineKeyboardButton(
                text=button_text,
                callback_data=f"view_kp_{kp_id}"
            )
        ])
    
    # Кнопка "Назад"
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")
    ])
    
    await callback.message.edit_text(
        text,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )


@router.callback_query(F.data.startswith("view_kp_"))
async def view_kp_details(callback: CallbackQuery):
    """
    Показать детали конкретного КП.
    
    Простыми словами:
    - Получает полную информацию о КП из БД
    - Показывает: клиента, менеджера, дату, сумму, статус, список плит
    - Показывает кнопки: Скачать PDF, Скачать XLSX, Удалить КП
    """
    await callback.answer()
    
    kp_id = int(callback.data.split("_")[-1])
    kp_info = kp_db.get_kp_by_id(kp_id)
    
    if not kp_info:
        await callback.message.edit_text(
            "❌ КП не найдено в базе данных",
            reply_markup=archive_sections_kb()
        )
        return
    
    await callback.message.edit_text(
        _format_kp_details_text(kp_info),
        reply_markup=kp_details_kb(kp_id, kp_info.get("total_amount") or 0, kp_info.get("status"))
    )


@router.callback_query(F.data.startswith("download_pdf_"))
async def download_pdf(callback: CallbackQuery):
    """
    Генерирует PDF КП на лету из данных БД (как при создании КП) и отправляет пользователю.
    """
    await callback.answer("⏳ Подготавливаю PDF...")
    kp_id = int(callback.data.split("_")[-1])
    kp_info = kp_db.get_kp_by_id(kp_id)
    if not kp_info:
        await callback.message.answer("❌ КП не найдено в базе данных")
        return
    order_data = _order_data_from_kp_info(kp_info)
    if not order_data:
        await callback.message.answer("❌ В КП нет позиций для формирования документа")
        return
    offer_number = str(kp_id)
    offer_date = kp_info.get("creation_date") or datetime.now().strftime("%d.%m.%Y")
    customer_name = kp_info.get("customer_name")
    manager_name = kp_info.get("manager_name")
    discount_percent = kp_info.get("discount_percent") or 0
    try:
        pdf_buffer = await asyncio.to_thread(
            generate_commercial_offer_pdf,
            order_data,
            offer_number,
            offer_date,
            customer_name=customer_name,
            manager_name=manager_name,
            manager_phone=None,
            manager_email=None,
            discount_percent=discount_percent,
            kp_db_id=kp_id,
        )
        await callback.message.answer_document(
            BufferedInputFile(pdf_buffer.getvalue(), filename=f"КП_{kp_id}.pdf"),
            caption=f"📄 КП № {kp_id} (PDF)",
        )
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка генерации PDF: {e}")


@router.callback_query(F.data.startswith("download_xlsx_"))
async def download_xlsx(callback: CallbackQuery):
    """
    Генерирует XLSX КП на лету из данных БД (как при создании КП) и отправляет пользователю.
    """
    await callback.answer("⏳ Подготавливаю XLSX...")
    kp_id = int(callback.data.split("_")[-1])
    kp_info = kp_db.get_kp_by_id(kp_id)
    if not kp_info:
        await callback.message.answer("❌ КП не найдено в базе данных")
        return
    order_data = _order_data_from_kp_info(kp_info)
    if not order_data:
        await callback.message.answer("❌ В КП нет позиций для формирования документа")
        return
    offer_number = str(kp_id)
    offer_date = kp_info.get("creation_date") or datetime.now().strftime("%d.%m.%Y")
    customer_name = kp_info.get("customer_name")
    manager_name = kp_info.get("manager_name")
    discount_percent = kp_info.get("discount_percent") or 0
    delivery_conditions = kp_info.get("delivery_conditions")
    payment_conditions = kp_info.get("payment_conditions")
    try:
        xlsx_buffer = await asyncio.to_thread(
            generate_commercial_offer_xlsx,
            order_data,
            offer_number,
            offer_date,
            customer_name=customer_name,
            manager_name=manager_name,
            manager_phone=None,
            manager_email=None,
            discount_percent=discount_percent,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            kp_db_id=kp_id,
        )
        await callback.message.answer_document(
            BufferedInputFile(xlsx_buffer.getvalue(), filename=f"КП_{kp_id}.xlsx"),
            caption=f"📊 КП № {kp_id} (XLSX с формулами)",
        )
    except Exception as e:
        await callback.message.answer(f"❌ Ошибка генерации XLSX: {e}")


def _format_kp_details_text(kp_info: dict) -> str:
    """Формирует текст карточки КП для отображения в деталях и после смены скидки."""
    kp_id = kp_info.get("kp_id", 0)
    text = f"📋 Коммерческое предложение № {kp_id}\n\n"
    text += f"👤 Клиент: {kp_info.get('customer_name', 'Не указан')}\n"
    text += f"👨‍💼 Менеджер: {kp_info.get('manager_name', 'Не указан')}\n"
    text += f"📅 Дата создания: {kp_info.get('creation_date', 'Не указана')}\n"
    status = kp_info.get("status", "Неизвестен")
    status_emoji = {"в архиве": "📦", "в работе": "🏭", "выполнено": "✅", "отклонено": "❌"}.get(status, "❓")
    text += f"📊 Статус: {status_emoji} {status}\n"
    if kp_info.get("execution_terms"):
        text += f"⏰ Срок выполнения: {kp_info['execution_terms']}\n"
    text += f"\n💰 Финансы:\n"
    text += f"  • Сумма без НДС: {kp_info.get('subtotal', 0):,.2f} ₽\n"
    text += f"  • НДС (22%): {kp_info.get('vat_amount', 0):,.2f} ₽\n"
    text += f"  • Итого с НДС: {kp_info.get('total_amount', 0):,.2f} ₽\n"
    if kp_info.get("discount_percent", 0) > 0:
        text += f"  • Скидка: {kp_info['discount_percent']}%\n"
    transport_hours = kp_info.get("transport_hours")
    transport_price_per_hour = kp_info.get("transport_price_per_hour")
    if transport_hours and transport_price_per_hour:
        transport_total = float(transport_hours) * float(transport_price_per_hour)
        text += (
            f"  • Транспорт: {format_offer_quantity(transport_hours)} {('час' if float(transport_hours) == 1 else 'час.')}"
            f" × {float(transport_price_per_hour):,.2f} ₽ = {transport_total:,.2f} ₽\n"
        )
    plates = kp_info.get("plates", [])
    text += f"\n📦 Состав заказа ({len(plates)} позиций):\n"
    for i, plate in enumerate(plates[:10], 1):
        text += f"  {i}. {plate.get('plate_name', '')} — {plate.get('qty', 0)} шт\n"
    if len(plates) > 10:
        text += f"  ... и ещё {len(plates) - 10} позиций\n"
    return text


@router.callback_query(F.data.startswith("transport_kp_"))
async def ask_transport_costs(callback: CallbackQuery, state: FSMContext):
    """Запрашивает количество часов транспортных расходов для выбранного КП."""
    await callback.answer()
    kp_id = int(callback.data.split("_")[-1])
    kp_info = kp_db.get_kp_by_id(kp_id)
    if not kp_info:
        await callback.message.answer("❌ КП не найдено в базе данных.")
        return

    await state.update_data(transport_kp_id=kp_id)
    await state.set_state(ArchiveStates.waiting_transport_hours)

    existing_hours = kp_info.get("transport_hours")
    existing_price = kp_info.get("transport_price_per_hour")
    existing_text = ""
    if existing_hours and existing_price:
        existing_text = (
            "\nТекущие значения:\n"
            f"• Часы: {format_offer_quantity(existing_hours)}\n"
            f"• Цена за час: {float(existing_price):,.2f} ₽\n"
        )

    await callback.message.answer(
        "🚚 Транспортные расходы\n\n"
        "Введите количество часов транспортных услуг."
        f"{existing_text}\n"
        "Пример: 2 или 2,5"
    )


@router.message(ArchiveStates.waiting_transport_hours, F.text)
async def receive_transport_hours(message: Message, state: FSMContext):
    """Сохраняет часы транспорта и переходит к запросу цены за час."""
    try:
        transport_hours = _parse_positive_number(message.text)
    except ValueError:
        await message.answer("Введите число больше 0. Например: 2 или 2,5.")
        return

    await state.update_data(transport_hours=transport_hours)
    await state.set_state(ArchiveStates.waiting_transport_price)
    await message.answer(
        f"✅ Часы транспорта: {format_offer_quantity(transport_hours)}\n\n"
        "Теперь введите цену транспортных услуг за 1 час."
    )


@router.message(ArchiveStates.waiting_transport_price, F.text)
async def receive_transport_price(message: Message, state: FSMContext):
    """Сохраняет транспортные расходы в КП и обновляет карточку."""
    data = await state.get_data()
    kp_id = data.get("transport_kp_id")
    transport_hours = data.get("transport_hours")
    if kp_id is None or transport_hours is None:
        await state.clear()
        await message.answer("Сессия сброшена. Откройте КП снова и повторите ввод.")
        return

    try:
        transport_price_per_hour = _parse_positive_number(message.text)
    except ValueError:
        await message.answer("Введите цену за час числом больше 0. Например: 9800 или 9800,50.")
        return

    success = kp_db.update_kp_transport(kp_id, transport_hours, transport_price_per_hour)
    await state.clear()
    if not success:
        await message.answer("❌ Не удалось сохранить транспортные расходы.")
        return

    kp_info = kp_db.get_kp_by_id(kp_id)
    if not kp_info:
        await message.answer("КП не найдено.")
        return

    text = _format_kp_details_text(kp_info)
    await message.answer(
        f"✅ Транспортные расходы сохранены.\n\n{text}",
        reply_markup=kp_details_kb(kp_id, kp_info.get("total_amount") or 0, kp_info.get("status")),
    )


@router.callback_query(F.data.startswith("change_discount_"))
async def ask_new_discount(callback: CallbackQuery, state: FSMContext):
    """Запрос нового процента скидки для КП."""
    await callback.answer()
    kp_id = int(callback.data.split("_")[-1])
    await state.update_data(discount_kp_id=kp_id)
    await state.set_state(ArchiveStates.waiting_discount)
    await callback.message.answer("Введите новый процент скидки (0–100):")


@router.message(ArchiveStates.waiting_discount, F.text)
async def apply_new_discount(message: Message, state: FSMContext):
    """Применяет введённый процент скидки и обновляет карточку КП."""
    data = await state.get_data()
    kp_id = data.get("discount_kp_id")
    if kp_id is None:
        await state.clear()
        await message.answer("Сессия сброшена. Выберите КП снова из архива.")
        return
    try:
        value = float(message.text.strip().replace(",", "."))
    except ValueError:
        await message.answer("Введите число от 0 до 100 (например: 5 или 10).")
        return
    if not (0 <= value <= 100):
        await message.answer("Процент скидки должен быть от 0 до 100.")
        return
    success = kp_db.update_kp_discount(kp_id, value)
    await state.clear()
    if not success:
        await message.answer("❌ Не удалось обновить скидку. Проверьте, что КП существует и в нём есть позиции.")
        return
    kp_info = kp_db.get_kp_by_id(kp_id)
    if not kp_info:
        await message.answer("КП не найдено.")
        return
    text = _format_kp_details_text(kp_info)
    await message.answer(
        f"✅ Скидка обновлена до {value}%.\n\n{text}",
        reply_markup=kp_details_kb(kp_id, kp_info.get("total_amount") or 0, kp_info.get("status")),
    )


@router.callback_query(F.data.startswith("delete_kp_"))
async def delete_kp_confirm(callback: CallbackQuery, state: FSMContext):
    """
    Подтверждение удаления КП.
    
    Простыми словами:
    - Показывает предупреждение об удалении
    - Просит подтвердить действие
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    kp_id = int(callback.data.split("_")[-1])
    
    # Получаем информацию о КП для показа в подтверждении
    kp_info = kp_db.get_kp_by_id(kp_id)
    customer = kp_info.get('customer_name', 'Неизвестен') if kp_info else 'Неизвестен'
    
    # Показываем подтверждение
    await callback.message.edit_text(
        f"⚠️ Вы уверены, что хотите удалить КП № {kp_id}?\n\n"
        f"Клиент: {customer}\n\n"
        f"Это действие НЕОБРАТИМО!\n"
        f"Будут удалены:\n"
        f"  • Информация о КП\n"
        f"  • Список плит\n"
        f"  • Файлы (PDF и XLSX)\n"
        f"  • Метаданные",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Да, удалить", callback_data=f"delete_kp_confirmed_{kp_id}")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data=f"view_kp_{kp_id}")]
        ])
    )


@router.callback_query(F.data.startswith("delete_kp_confirmed_"))
async def delete_kp_execute(callback: CallbackQuery):
    """
    Выполнить удаление КП из базы данных.
    
    Простыми словами:
    - Удаляет КП из БД (все связанные записи удаляются автоматически)
    - Показывает результат операции
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    kp_id = int(callback.data.split("_")[-1])
    
    # Выполняем удаление
    success = kp_db.delete_kp_by_id(kp_id)
    
    if success:
        await callback.message.edit_text(
            f"✅ КП № {kp_id} успешно удалён из базы данных\n\n"
            f"Удалены:\n"
            f"  • Основная информация о КП\n"
            f"  • Все позиции (плиты)\n"
            f"  • Файлы из БД\n"
            f"  • Метаданные и статус",
            reply_markup=archive_sections_kb()
        )
    else:
        await callback.message.edit_text(
            f"❌ Ошибка при удалении КП № {kp_id}\n\n"
            f"Возможно, КП уже был удалён ранее.",
            reply_markup=archive_sections_kb()
        )


@router.callback_query(F.data.startswith("move_kp_to_production_"))
async def move_kp_to_production(callback: CallbackQuery, state: FSMContext):
    """
    Переводит КП из архива в производство (статус «в работе»).
    Сначала запрашивает срок выполнения (как «Сохранить в БД»), сохраняет в БД, затем меняет статус.
    Кнопка доступна только для КП со статусом «в архиве».
    """
    try:
        await callback.answer()
    except Exception:
        pass
    kp_id = int(callback.data.split("_")[-1])
    kp_info = kp_db.get_kp_by_id(kp_id)
    if not kp_info:
        await callback.message.edit_text(
            "❌ КП не найдено в базе данных",
            reply_markup=archive_sections_kb()
        )
        return
    if kp_info.get("status") != "в архиве":
        text = _format_kp_details_text(kp_info)
        await callback.message.edit_text(
            text,
            reply_markup=kp_details_kb(
                kp_id,
                kp_info.get("total_amount") or 0,
                kp_info.get("status"),
            )
        )
        return
    # Оценка производства из plates (те же формулы, что в commercial.py при «Сохранить в БД»)
    MAX_TRACK_LENGTH = 101.0
    plates = kp_info.get("plates") or []
    total_length = sum(p.get("length_m", 0) * p.get("qty", 1) for p in plates)
    estimated_tracks = max(1, int(round(total_length / MAX_TRACK_LENGTH + 0.5)))
    estimated_days = max(1, int(round(estimated_tracks / 5.0 + 0.5)))
    await state.update_data(production_kp_id=kp_id)
    await state.set_state(ArchiveStates.waiting_execution_terms_for_production)
    await callback.message.edit_text(
        "✅ Перевожу КП в производство...\n\n"
        f"⏱️ Оценка производства:\n"
        f"  • Примерно дорожек: {estimated_tracks}\n"
        f"  • Примерно дней: {estimated_days}\n\n"
        "📅 Укажите сроки выполнения:\n"
        f"(Например: '{estimated_days} дней', '2 недели', '01.02.2024')"
    )


@router.message(ArchiveStates.waiting_execution_terms_for_production)
async def receive_production_execution_terms(message: Message, state: FSMContext):
    """
    Обработчик ввода сроков выполнения при переводе КП из архива в производство.
    Парсит срок (дата / N дней / N недель), сохраняет в БД, переводит статус в «в работе».
    """
    execution_terms_input = message.text.strip()
    logger.debug(f"[Архив → производство] Получены сроки: {execution_terms_input}")
    data = await state.get_data()
    kp_id = data.get("production_kp_id")
    if not kp_id:
        await message.answer(
            "❌ Сессия истекла. Выберите КП в архиве и нажмите «В производство» снова.",
            reply_markup=archive_sections_kb()
        )
        await state.clear()
        return
    deadline_date = None
    # Вариант 1: ДД.ММ.ГГГГ
    try:
        deadline_date = datetime.strptime(execution_terms_input, "%d.%m.%Y")
    except ValueError:
        pass
    # Вариант 2: ГГГГ-ММ-ДД
    if not deadline_date:
        try:
            deadline_date = datetime.strptime(execution_terms_input, "%Y-%m-%d")
        except ValueError:
            pass
    # Вариант 3: N дней
    if not deadline_date:
        match_days = re.search(r"(\d+)\s*(?:дн|день|дней|day|days)", execution_terms_input, re.IGNORECASE)
        if match_days:
            days = int(match_days.group(1))
            deadline_date = datetime.now() + timedelta(days=days)
    # Вариант 4: N недель
    if not deadline_date:
        match_weeks = re.search(r"(\d+)\s*(?:нед|недел|недели|week|weeks)", execution_terms_input, re.IGNORECASE)
        if match_weeks:
            weeks = int(match_weeks.group(1))
            deadline_date = datetime.now() + timedelta(weeks=weeks)
    if not deadline_date:
        deadline_date = datetime.now() + timedelta(days=14)
        await message.answer(
            f"⚠️ Не удалось распознать формат срока.\n"
            f"Использую значение по умолчанию: 14 дней\n"
            f"Дедлайн: {deadline_date.strftime('%d.%m.%Y')}"
        )
    execution_terms = deadline_date.strftime("%d.%m.%Y")
    try:
        ok_date = kp_db.update_kp_execution_date(kp_id, execution_terms)
        if not ok_date:
            await message.answer(
                f"❌ Не удалось обновить срок для КП № {kp_id}.",
                reply_markup=archive_sections_kb()
            )
            await state.clear()
            return
        success = kp_db.update_kp_status(kp_id, "в работе")
        if not success:
            await message.answer(
                f"❌ Срок сохранён, но не удалось перевести КП № {kp_id} в производство.",
                reply_markup=archive_sections_kb()
            )
            await state.clear()
            return
        kp_info = kp_db.get_kp_by_id(kp_id)
        customer_name = (kp_info.get("customer_name") or "—") if kp_info else "—"
        await message.answer(
            f"✅ КП № {kp_id} переведено в производство.\n\n"
            f"  • Клиент: {customer_name}\n"
            f"  • Срок изготовления до: {execution_terms}\n"
            f"  • Статус: в работе",
            reply_markup=archive_sections_kb()
        )
    except Exception as e:
        logger.exception(f"Ошибка при переводе КП {kp_id} в производство: {e}")
        await message.answer(
            "❌ Ошибка при сохранении. Попробуйте позже.",
            reply_markup=archive_sections_kb()
        )
    finally:
        await state.clear()


@router.callback_query(F.data == "view_current_plan")
async def view_current_plan(callback: CallbackQuery):
    """
    Просмотр актуального плана производства.
    Строит суммарную диаграмму Ганта по всем сохранённым планам (как кнопка «Диаграмма Ганта» в производстве).
    """
    try:
        await callback.answer()
    except Exception:
        pass
    await callback.message.answer("📊 Создаю суммарную диаграмму Ганта по всем планам...")
    gantt_data = get_all_plans_gantt_data()
    if not gantt_data:
        await callback.message.answer(
            "❌ Нет сохранённых планов для создания диаграммы.\n\n"
            "💡 Сначала создайте и сохраните план:\n"
            "1️⃣ Нажмите «🚀 Начать планирование»\n"
            "2️⃣ Выберите КП для производства\n"
            "3️⃣ Нажмите «💾 Сохранить план»",
            reply_markup=archive_sections_kb()
        )
        return
    all_tracks_list = gantt_data["all_tracks"]
    plate_lookup_exact = gantt_data["plate_lookup_exact"]
    plate_lookup_by_length = gantt_data["plate_lookup_by_length"]
    start_date_for_gantt = gantt_data["earliest_start_date"]
    plans_count = gantt_data["plans_count"]
    total_days = gantt_data["total_days"]
    tracks_count = 3
    try:
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_list,
            tracks_count=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            output_dir=OUTPUTS_DIR_STR,
            start_date=start_date_for_gantt,
        )
        if gantt_path and os.path.exists(gantt_path):
            start_date_str = start_date_for_gantt.strftime("%d.%m.%Y")
            end_date_str = gantt_data["latest_end_date"].strftime("%d.%m.%Y")
            await callback.message.answer_document(
                FSInputFile(gantt_path),
                caption=(
                    "📊 СУММАРНАЯ диаграмма Ганта по всем планам\n\n"
                    f"📅 Период: {start_date_str} — {end_date_str}\n"
                    f"📋 Планов: {plans_count}\n"
                    f"📆 Дней: {total_days}\n"
                    f"🛤️ Дорожек: {len(all_tracks_list)}\n\n"
                    "Цветовая кодировка:\n"
                    "🟢 Зелёный — успеваем до дедлайна\n"
                    "🟡 Жёлтый — завершаем в день дедлайна\n"
                    "🔴 Красный — опаздываем!"
                ),
            )
            logger.info("[GANTT] Диаграмма успешно создана из архива: %s", gantt_path)
        else:
            await callback.message.answer(
                "⚠️ Не удалось создать диаграмму.\n"
                "Возможно, нет данных о КП в сохранённых планах.\n\n"
                "💡 Убедитесь, что планы содержат информацию о заказах.",
                reply_markup=archive_sections_kb(),
            )
            return
    except Exception as e:
        logger.exception("Ошибка создания диаграммы из архива: %s", e)
        await callback.message.answer(
            "❌ Не удалось создать диаграмму Ганта.\nПодробности в logs/bot.log.",
            reply_markup=archive_sections_kb(),
        )
        return
    await callback.message.answer(
        "Выберите раздел для просмотра:",
        reply_markup=archive_sections_kb(),
    )


@router.callback_query(F.data == "archive_back_to_sections")
async def back_to_sections(callback: CallbackQuery, state: FSMContext):
    """Вернуться к выбору раздела архива"""
    try:
        await callback.answer()
    except Exception:
        pass  # Игнорируем ошибку, если callback устарел
    await state.clear()
    await callback.message.edit_text(
        "📁 Архив коммерческих предложений\n\n"
        "Выберите раздел для просмотра:",
        reply_markup=archive_sections_kb()
    )


@router.callback_query(F.data == "archive_back_to_menu")
async def back_to_menu(callback: CallbackQuery):
    """Вернуться в главное меню"""
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    await callback.message.answer(
        "Главное меню:",
        reply_markup=main_menu_kb()
    )
