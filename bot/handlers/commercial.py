"""Обработчики создания коммерческих предложений PDF/XLSX"""
import asyncio
import os
import re
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, Any

from aiogram import Router, F
from aiogram.types import Message, FSInputFile, CallbackQuery
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext

# Добавляем корень проекта в sys.path
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.config_and_data import set_plate_lists_from_text
import core.config_and_data as cfg
from core.commercial_offer import generate_commercial_offer_pdf
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.visualization import visualize_plan
from core import kp_db

from ..keyboards import main_menu_kb, conditions_choice_kb, save_to_db_kb
from ..states import KPStates
from ..bot_config import OUTPUTS_DIR_STR

router = Router()

# Кэш заказов пользователей
ORDER_CACHE: Dict[int, list] = {}


@router.message(F.text == "Коммерческое предложение PDF")
@router.message(Command("commercial_offer"))
async def btn_commercial_offer(message: Message, state: FSMContext):
    """Обработчик запроса на создание коммерческого предложения - ПОШАГОВЫЙ ОПРОС"""
    # region agent log
    import json
    with open('/home/username/Рабочий стол/my py/plan_AI/.cursor/debug.log', 'a') as f:
        f.write(json.dumps({"location":"commercial.py:38","message":"Starting step 1 - manager name","data":{"user_id":message.from_user.id},"timestamp":__import__('time').time()*1000,"sessionId":"debug-session","runId":"run1","hypothesisId":"A"})+'\n')
    # endregion
    # Шаг 1: Запрашиваем имя менеджера
    await state.set_state(KPStates.waiting_manager_name)
    await message.answer(
        "📄 Создание коммерческого предложения\n\n"
        "Шаг 1 из 5: Введите имя менеджера\n"
        "(Это имя будет указано в документе)",
        reply_markup=main_menu_kb()
    )


# === ПОШАГОВЫЙ ОПРОС ДЛЯ КОММЕРЧЕСКОГО ПРЕДЛОЖЕНИЯ ===

@router.message(KPStates.waiting_manager_name)
async def receive_manager_name(message: Message, state: FSMContext):
    """Шаг 1: Получаем имя менеджера"""
    manager_name = message.text.strip()
    
    # region agent log
    import json
    with open('/home/username/Рабочий стол/my py/plan_AI/.cursor/debug.log', 'a') as f:
        f.write(json.dumps({"location":"commercial.py:53","message":"Received manager name, moving to step 2","data":{"manager_name":manager_name},"timestamp":__import__('time').time()*1000,"sessionId":"debug-session","runId":"run1","hypothesisId":"B"})+'\n')
    # endregion
    
    # Сохраняем имя менеджера в состояние
    await state.update_data(manager_name=manager_name)
    
    # Переходим к следующему шагу - запрашиваем имя клиента
    await state.set_state(KPStates.waiting_client_name)
    await message.answer(
        f"✅ Менеджер: {manager_name}\n\n"
        "Шаг 2 из 5: Введите имя клиента\n"
        "(Для кого создается коммерческое предложение)",
        reply_markup=main_menu_kb()
    )


@router.message(KPStates.waiting_client_name)
async def receive_client_name(message: Message, state: FSMContext):
    """Шаг 2: Получаем имя клиента"""
    client_name = message.text.strip()
    
    # region agent log
    import json
    with open('/home/username/Рабочий стол/my py/plan_AI/.cursor/debug.log', 'a') as f:
        f.write(json.dumps({"location":"commercial.py:79","message":"Received client name, moving to step 3","data":{"client_name":client_name},"timestamp":__import__('time').time()*1000,"sessionId":"debug-session","runId":"run1","hypothesisId":"C"})+'\n')
    # endregion
    
    # Сохраняем имя клиента в состояние
    await state.update_data(client_name=client_name)
    
    # Переходим к следующему шагу - запрашиваем список плит
    await state.set_state(KPStates.waiting_plates_list)
    await message.answer(
        f"✅ Клиент: {client_name}\n\n"
        "Шаг 3 из 5: Пришлите список плит в свободной форме\n\n"
        "Примеры форматов:\n"
        "• '1.2×3.39 — 2 шт'\n"
        "• '0.32×6.63 — 4 шт'\n"
        "• 'ПБ 38-12-8п 2'\n"
        "• 'ПБ 66-3-8п 4'",
        reply_markup=main_menu_kb()
    )


@router.message(KPStates.waiting_plates_list)
async def receive_plates_list(message: Message, state: FSMContext):
    """Шаг 3: Получаем список плит"""
    plates_text = message.text.strip()
    
    # Сохраняем список плит в состояние
    await state.update_data(plates_text=plates_text)
    
    # Переходим к следующему шагу - запрашиваем процент скидки
    await state.set_state(KPStates.waiting_discount)
    await message.answer(
        "✅ Список плит получен\n\n"
        "Шаг 4 из 5: Введите процент скидки\n"
        "(Просто число, например: 0, 5, 10, 15)\n"
        "0 = без скидки",
        reply_markup=main_menu_kb()
    )


@router.message(KPStates.waiting_discount)
async def receive_discount_and_ask_conditions(message: Message, state: FSMContext):
    """Шаг 4: Получаем процент скидки и переходим к условиям"""
    try:
        # Парсим процент скидки
        discount_text = message.text.strip().replace('%', '').replace(',', '.')
        discount_percent = float(discount_text)
        
        if discount_percent < 0 or discount_percent > 100:
            await message.answer(
                "❌ Процент скидки должен быть от 0 до 100\n"
                "Попробуйте снова:",
                reply_markup=main_menu_kb()
            )
            return
    except ValueError:
        await message.answer(
            "❌ Неверный формат числа. Введите просто число (например: 0, 5, 10):",
            reply_markup=main_menu_kb()
        )
        return
    
    # Сохраняем скидку в состояние
    await state.update_data(discount_percent=discount_percent)
    
    # Переходим к шагу 5 - выбор условий
    await state.set_state(KPStates.waiting_conditions_choice)
    await message.answer(
        f"✅ Скидка: {discount_percent}%\n\n"
        "Шаг 5 из 5: Условия поставки и оплаты\n\n"
        "Выберите вариант:",
        reply_markup=conditions_choice_kb()
    )


@router.callback_query(KPStates.waiting_conditions_choice)
async def receive_conditions_choice(callback: CallbackQuery, state: FSMContext):
    """Шаг 5: Выбор условий (по умолчанию или добавить свои) - обработка inline-кнопок"""
    choice = callback.data
    
    # Убираем "часики" с кнопки
    await callback.answer()
    
    if choice == "conditions_default":
        # Используем значения по умолчанию
        await state.update_data(
            delivery_conditions="",  # Пустая строка = использовать по умолчанию
            payment_conditions=""    # Пустая строка = использовать по умолчанию
        )
        
        # Редактируем сообщение с кнопками
        await callback.message.edit_text(
            "✅ Выбрано: По умолчанию"
        )
        
        # Переходим сразу к генерации
        await generate_all_documents(callback.message, state)
        
    elif choice == "conditions_custom":
        # Редактируем сообщение с кнопками
        await callback.message.edit_text(
            "✅ Выбрано: Добавить условие"
        )
        
        # Запрашиваем условия поставки
        await state.set_state(KPStates.waiting_delivery_conditions)
        await callback.message.answer(
            "Введите условия поставки:\n"
            "(Например: 'Самовывоз со склада' или 'Доставка до объекта')",
            reply_markup=main_menu_kb()
        )


@router.message(KPStates.waiting_delivery_conditions)
async def receive_delivery_conditions(message: Message, state: FSMContext):
    """Получаем условия поставки"""
    delivery_conditions = message.text.strip()
    
    # Сохраняем условия поставки
    await state.update_data(delivery_conditions=delivery_conditions)
    
    # Переходим к запросу условий оплаты
    await state.set_state(KPStates.waiting_payment_conditions)
    await message.answer(
        f"✅ Условия поставки: {delivery_conditions}\n\n"
        "Теперь введите условия оплаты:\n"
        "(Например: 'Предварительная оплата 100%' или '50% аванс, 50% по факту отгрузки')",
        reply_markup=main_menu_kb()
    )


@router.message(KPStates.waiting_payment_conditions)
async def receive_payment_conditions_and_generate(message: Message, state: FSMContext):
    """Получаем условия оплаты и генерируем документы"""
    payment_conditions = message.text.strip()
    
    # Сохраняем условия оплаты
    await state.update_data(payment_conditions=payment_conditions)
    
    # Переходим к генерации документов
    await generate_all_documents(message, state)


async def generate_all_documents(message: Message, state: FSMContext):
    """Генерация всех документов с полученными данными"""
    # Получаем все данные из состояния
    data = await state.get_data()
    manager_name = data.get('manager_name', 'Не указано')
    client_name = data.get('client_name', 'Не указано')
    plates_text = data.get('plates_text', '')
    discount_percent = data.get('discount_percent', 0)
    delivery_conditions = data.get('delivery_conditions', '')
    payment_conditions = data.get('payment_conditions', '')
    
    # Показываем сводку
    summary_text = (
        f"📋 Сводка:\n"
        f"• Менеджер: {manager_name}\n"
        f"• Клиент: {client_name}\n"
        f"• Скидка: {discount_percent}%\n"
    )
    if delivery_conditions:
        summary_text += f"• Условия поставки: {delivery_conditions}\n"
    if payment_conditions:
        summary_text += f"• Условия оплаты: {payment_conditions}\n"
    
    summary_text += f"\n⏳ Формирую коммерческое предложение..."
    await message.answer(summary_text)
    
    # Теперь запускаем генерацию документов с этими данными
    try:
        # Парсим список пользователя
        unparsed_lines = set_plate_lists_from_text(plates_text)
        
        if unparsed_lines:
            warn_text = "⚠️ Некоторые строки я не смог распознать по формату и пропустил:\n"
            warn_text += "\n".join(f"• {line}" for line in unparsed_lines)
            warn_text += (
                "\n\nЯ понимаю, например, такие форматы:\n"
                "• 1.2×3.39 — 2 шт\n"
                "• 0,32x6,63 - 4\n"
                "• Плиты ПБ 78-12-8п 3\n"
                "• ПБ 66,2-12-8п 6\n"
            )
            await message.answer(warn_text)
        
        # Используем build_price_rows для получения правильных цен
        from viz_modules.procurement import build_price_rows
        from viz_modules.price_utils import load_price_table_from_xlsx
        
        # Загружаем таблицу цен
        price_table = load_price_table_from_xlsx(cfg.PRICE_XLSX_PATH)
        
        # Получаем строки сметы
        price_rows, total_sum = await asyncio.to_thread(
            build_price_rows,
            price_table,
            reinforcement_code=8
        )
        
        # Формируем order_data
        order_data = []
        for row in price_rows:
            if len(row) < 8:
                continue
            
            name = row[1]
            qty = row[2]
            weight_str = row[6]
            price_str = row[7]
            
            try:
                unit_price = float(price_str.replace(' ', '').replace(',', '.'))
            except (ValueError, AttributeError):
                unit_price = 0.0
            
            try:
                weight = float(str(weight_str).replace(' ', '').replace(',', '.'))
            except (ValueError, AttributeError):
                weight = 0.0
            
            # Парсим name
            match = re.search(r'ПБ\s+(\d+)-([\d,]+)-(\d+)', name)
            if match:
                length_dm = int(match.group(1))
                length_m = length_dm / 10.0
                
                width_str_parsed = match.group(2).replace(',', '.')
                width_m = float(width_str_parsed)
                if width_m > 2:
                    width_m = width_m / 10.0
                
                load_code = int(match.group(3))
                
                order_data.append({
                    "name": name,
                    "length_m": length_m,
                    "width_m": width_m,
                    "qty": qty,
                    "load_class": load_code * 100,
                    "unit_price": unit_price,
                    "weight": weight
                })
        
        if not order_data:
            await message.answer(
                "❌ Не удалось распознать ни одной плиты в вашем сообщении.\n"
                "Проверьте формат строк.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        # Сохраняем заказ в кэш
        ORDER_CACHE[message.from_user.id] = order_data
        
        # Генерируем номер и дату КП
        offer_number = f"{message.from_user.id}_{datetime.now().strftime('%Y%m%d%H%M')}"
        offer_date = datetime.now().strftime("%d.%m.%Y")
        
        # Генерируем PDF (старая версия, не использует имя менеджера)
        pdf_buffer = await asyncio.to_thread(
            generate_commercial_offer_pdf,
            order_data,
            offer_number,
            offer_date,
            client_name  # Используем имя клиента из опроса
        )
        
        # Генерируем XLSX с НОВЫМИ параметрами (менеджер, скидка, условия)
        try:
            xlsx_buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data,
                offer_number,
                offer_date,
                client_name,         # Используем имя клиента
                manager_name,        # НОВОЕ: передаем имя менеджера
                discount_percent,    # НОВОЕ: передаем процент скидки
                delivery_conditions, # НОВОЕ: условия поставки
                payment_conditions   # НОВОЕ: условия оплаты
            )
            has_xlsx = True
        except Exception as e:
            print(f"[XLSX] Ошибка генерации XLSX: {e}")
            import traceback
            traceback.print_exc()
            has_xlsx = False
        
        # Сохраняем файлы
        pdf_filename = f"КП_{offer_number}_{offer_date.replace('.', '')}.pdf"
        pdf_path = os.path.join(OUTPUTS_DIR_STR, pdf_filename)
        
        xlsx_filename = f"КП_{offer_number}_{offer_date.replace('.', '')}.xlsx"
        xlsx_path = os.path.join(OUTPUTS_DIR_STR, xlsx_filename)
        
        with open(pdf_path, 'wb') as f:
            f.write(pdf_buffer.getvalue())
        
        if has_xlsx:
            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_buffer.getvalue())
        
        # Формируем сводку
        total_qty = sum(item['qty'] for item in order_data)
        summary = f"✅ Коммерческое предложение готово!\n\n"
        summary += f"📋 Заказ:\n"
        for item in order_data:
            summary += f"  • {item['name']} — {item['qty']} шт\n"
        summary += f"\n📊 Всего позиций: {len(order_data)}\n"
        summary += f"📦 Всего плит: {total_qty} шт\n"
        
        await message.answer(summary)
        
        # Отправляем PDF
        if os.path.exists(pdf_path):
            await message.answer_document(
                FSInputFile(pdf_path),
                caption=f"📄 Коммерческое предложение № {offer_number} (PDF)"
            )
        
        # Отправляем XLSX
        if has_xlsx and os.path.exists(xlsx_path):
            await message.answer_document(
                FSInputFile(xlsx_path),
                caption=f"📊 Коммерческое предложение № {offer_number} (XLSX с формулами)"
            )
        
        # Генерируем схему и разбивку
        await message.answer("⏳ Генерирую схему раскладки и детальную разбивку...")
        
        try:
            result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR_STR)
            
            if isinstance(result_paths, tuple) and len(result_paths) >= 2:
                png_path, pdf_schema_path = result_paths
                
                base = os.path.basename(png_path)
                if 'КЗ_' in base:
                    timestamp = base.split('КЗ_', 1)[-1].replace('.png', '')
                else:
                    timestamp = base.rsplit('_', 1)[-1].replace('.png', '')
                
                breakdown_path = os.path.join(OUTPUTS_DIR_STR, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx')
                
                if os.path.exists(pdf_schema_path):
                    await message.answer_document(
                        FSInputFile(pdf_schema_path),
                        caption="📐 Схема раскладки плит"
                    )
                
                if os.path.exists(breakdown_path):
                    await message.answer_document(
                        FSInputFile(breakdown_path),
                        caption="📊 Детальная разбивка компонентов"
                    )
        except Exception as e:
            print(f"[ОШИБКА] При генерации схемы: {e}")
            import traceback
            traceback.print_exc()
        
        await message.answer(
            "✨ Документы содержат:\n"
            "• Подробную спецификацию\n"
            "• Расчёт стоимости материалов\n"
            "• Стоимость резов\n"
            "• Вес изделий\n"
            "• НДС (20%)\n"
            "• Условия оплаты\n\n"
            f"💰 Скидка: {discount_percent}%\n"
            f"👤 Менеджер: {manager_name}\n"
            f"📊 XLSX файл содержит расчётные формулы Excel!"
        )
        
        # Сохраняем данные КП для последующего сохранения в БД
        # ВАЖНО: Сохраняем ДО очистки состояния
        await state.update_data(
            kp_order_data=order_data,
            kp_xlsx_path=xlsx_path if has_xlsx and os.path.exists(xlsx_path) else None,
            kp_customer_name=client_name,
            kp_manager_name=manager_name,
            kp_discount_percent=discount_percent,
            kp_delivery_conditions=delivery_conditions,
            kp_payment_conditions=payment_conditions,
            kp_offer_date=offer_date
        )
        
        # Очищаем состояние, чтобы callback мог сработать
        # Но данные остаются в state.data
        await state.set_state(None)
        
        print(f"[DEBUG] Данные сохранены в state:")
        print(f"  - Клиент: {client_name}")
        print(f"  - Менеджер: {manager_name}")
        print(f"  - Скидка: {discount_percent}%")
        print(f"  - Плит: {len(order_data)}")
        
        # Предлагаем сохранить КП в базу данных
        await message.answer(
            "💾 Хотите сохранить это КП в базу данных?\n\n"
            "Если сохраните, вы сможете отслеживать статус выполнения заказа.",
            reply_markup=save_to_db_kb()
        )
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при генерации КП: {str(e)}\n\n"
            "Проверьте формат данных и попробуйте снова.",
            reply_markup=main_menu_kb()
        )
        import traceback
        traceback.print_exc()
        await state.clear()


# === СТАРЫЙ ОБРАБОТЧИК (для обратной совместимости) ===

@router.message(KPStates.waiting_for_commercial_offer)
async def receive_order_and_generate_pdf(message: Message, state: FSMContext):
    """Обработчик получения заказа и генерации PDF"""
    await message.answer("⏳ Формирую коммерческое предложение...")
    
    try:
        # Парсим список пользователя и собираем «странные» строки по формату
        user_text = message.text or ""
        unparsed_lines = set_plate_lists_from_text(user_text)

        if unparsed_lines:
            warn_text = "⚠️ Некоторые строки я не смог распознать по формату и пропустил:\n"
            warn_text += "\n".join(f"• {line}" for line in unparsed_lines)
            warn_text += (
                "\n\nЯ понимаю, например, такие форматы:\n"
                "• 1.2×3.39 — 2 шт\n"
                "• 0,32x6,63 - 4\n"
                "• Плиты ПБ 78-12-8п 3\n"
                "• ПБ 66,2-12-8п 6\n"
            )
            await message.answer(warn_text)
        
        # 🔥 НОВАЯ ЛОГИКА: Используем build_price_rows для получения ПРАВИЛЬНЫХ цен
        # (с учётом резов, отходов и остатков - как в смете "Получить КП")
        from viz_modules.procurement import build_price_rows
        from viz_modules.price_utils import load_price_table_from_xlsx
        
        # Загружаем таблицу цен для расчётов
        price_table = load_price_table_from_xlsx(cfg.PRICE_XLSX_PATH)
        
        # Получаем строки сметы с ПРАВИЛЬНЫМИ ценами (как в "Получить КП")
        price_rows, total_sum = await asyncio.to_thread(
            build_price_rows,
            price_table,
            reinforcement_code=8
        )
        
        # Формируем order_data из price_rows с правильными ценами
        order_data = []
        for row in price_rows:
            # row: [idx, name, qty, 'шт', week, contractor, weight, price_str, sum_str]
            if len(row) < 8:
                continue
                
            name = row[1]
            qty = row[2]
            weight_str = row[6]
            price_str = row[7]  # Строка вида "1 234,56"
            
            # Парсим цену обратно в число
            try:
                unit_price = float(price_str.replace(' ', '').replace(',', '.'))
            except (ValueError, AttributeError):
                unit_price = 0.0
            
            # Парсим weight
            try:
                weight = float(str(weight_str).replace(' ', '').replace(',', '.'))
            except (ValueError, AttributeError):
                weight = 0.0
            
            # Парсим name для получения длины, ширины и нагрузки
            # Формат: "Плиты ПБ 38-12-8п" или "ПБ 38-3,2-8п"
            match = re.search(r'ПБ\s+(\d+)-([\d,]+)-(\d+)', name)
            if match:
                length_dm = int(match.group(1))
                length_m = length_dm / 10.0
                
                width_str_parsed = match.group(2).replace(',', '.')
                width_m = float(width_str_parsed)
                # Если ширина больше 2, значит она указана в дециметрах (например, 12 вместо 1.2)
                if width_m > 2:
                    width_m = width_m / 10.0
                
                load_code = int(match.group(3))
                
                order_data.append({
                    "name": name,
                    "length_m": length_m,
                    "width_m": width_m,
                    "qty": qty,
                    "load_class": load_code * 100,  # 8 → 800
                    "unit_price": unit_price,  # 🔥 Цена С УЧЁТОМ РЕЗОВ, ОТХОДОВ И ОСТАТКОВ!
                    "weight": weight
                })
        
        if not order_data:
            await message.answer(
                "❌ Не удалось распознать ни одной плиты в вашем сообщении.\n"
                "Проверьте формат строк (ширина×длина×кол-во или 'Плиты ПБ 78-12-8п 3').",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        # Сохраняем заказ в кэш
        ORDER_CACHE[message.from_user.id] = order_data
        
        # Генерируем номер и дату КП
        offer_number = f"{message.from_user.id}_{datetime.now().strftime('%Y%m%d%H%M')}"
        offer_date = datetime.now().strftime("%d.%m.%Y")
        
        # Получаем имя пользователя для КП
        user = message.from_user
        if user.last_name:
            customer_name = f"{user.first_name} {user.last_name}"
        else:
            customer_name = user.first_name or "заказчик"
        
        # Генерируем PDF и XLSX в памяти
        pdf_buffer = await asyncio.to_thread(
            generate_commercial_offer_pdf,
            order_data,
            offer_number,
            offer_date,
            customer_name
        )
        
        # Генерируем XLSX
        try:
            xlsx_buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data,
                offer_number,
                offer_date,
                customer_name
            )
            has_xlsx = True
        except Exception as e:
            print(f"[XLSX] Ошибка генерации XLSX: {e}")
            has_xlsx = False
        
        # Сохраняем файлы во временные файлы для отправки
        pdf_filename = f"КП_{offer_number}_{offer_date.replace('.', '')}.pdf"
        pdf_path = os.path.join(OUTPUTS_DIR_STR, pdf_filename)
        
        xlsx_filename = f"КП_{offer_number}_{offer_date.replace('.', '')}.xlsx"
        xlsx_path = os.path.join(OUTPUTS_DIR_STR, xlsx_filename)
        
        with open(pdf_path, 'wb') as f:
            f.write(pdf_buffer.getvalue())
        
        if has_xlsx:
            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_buffer.getvalue())
        
        # Формируем сводку по заказу
        total_qty = sum(item['qty'] for item in order_data)
        summary = f"✅ Коммерческое предложение готово!\n\n"
        summary += f"📋 Заказ:\n"
        for item in order_data:
            summary += f"  • {item['name']} — {item['qty']} шт\n"
        summary += f"\n📊 Всего позиций: {len(order_data)}\n"
        summary += f"📦 Всего плит: {total_qty} шт\n"
        
        await message.answer(summary)
        
        # Отправляем PDF
        if os.path.exists(pdf_path):
            await message.answer_document(
                FSInputFile(pdf_path),
                caption=f"📄 Коммерческое предложение № {offer_number} (PDF)"
            )
        
        # Отправляем XLSX
        if has_xlsx and os.path.exists(xlsx_path):
            await message.answer_document(
                FSInputFile(xlsx_path),
                caption=f"📊 Коммерческое предложение № {offer_number} (XLSX с формулами)"
            )
        
        # 🔥 ГЕНЕРИРУЕМ СХЕМУ И ДЕТАЛЬНУЮ РАЗБИВКУ (как при "Получить КП")
        await message.answer("⏳ Генерирую схему раскладки и детальную разбивку...")
        
        try:
            # Запускаем визуализацию (создаёт PDF схемы и XLSX детальной разбивки)
            result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR_STR)
            
            if isinstance(result_paths, tuple) and len(result_paths) >= 2:
                png_path, pdf_schema_path = result_paths
                
                # Извлекаем timestamp из имени PNG для поиска дополнительных файлов
                base = os.path.basename(png_path)
                if 'КЗ_' in base:
                    timestamp = base.split('КЗ_', 1)[-1].replace('.png', '')
                else:
                    timestamp = base.rsplit('_', 1)[-1].replace('.png', '')
                
                # Ищем файл детальной разбивки
                breakdown_path = os.path.join(OUTPUTS_DIR_STR, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx')
                
                # Отправляем PDF схемы
                if os.path.exists(pdf_schema_path):
                    await message.answer_document(
                        FSInputFile(pdf_schema_path),
                        caption="📐 Схема раскладки плит"
                    )
                
                # Отправляем XLSX детальной разбивки
                if os.path.exists(breakdown_path):
                    await message.answer_document(
                        FSInputFile(breakdown_path),
                        caption="📊 Детальная разбивка компонентов"
                    )
        except Exception as e:
            print(f"[ОШИБКА] При генерации схемы и разбивки: {e}")
            import traceback
            traceback.print_exc()
            # Не прерываем выполнение, просто пропускаем эти файлы
            
        if os.path.exists(pdf_path):
            await message.answer(
                "✨ Документы содержат:\n"
                "• Подробную спецификацию\n"
                "• Расчёт стоимости материалов\n"
                "• Стоимость резов\n"
                "• Вес изделий\n"
                "• НДС (20%)\n"
                "• Условия оплаты\n\n"
                "📊 XLSX файл содержит расчётные формулы Excel!\n"
                "📐 Схема раскладки и детальная разбивка также готовы!",
                reply_markup=main_menu_kb()
            )
        else:
            await message.answer(
                "❌ Ошибка при сохранении файла",
                reply_markup=main_menu_kb()
            )
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при генерации КП: {str(e)}\n\n"
            "Проверьте формат данных и попробуйте снова.",
            reply_markup=main_menu_kb()
        )
    finally:
        await state.clear()


# ==================== ОБРАБОТЧИКИ СОХРАНЕНИЯ КП В БД ====================

@router.callback_query(F.data == "save_kp_to_db")
async def callback_save_kp_to_db(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "Сохранить в БД".
    Запрашивает сроки выполнения КП.
    """
    print("[DEBUG] Нажата кнопка 'Сохранить в БД'")
    
    # Убираем "часики" с кнопки
    await callback.answer()
    
    # Редактируем сообщение с кнопками
    await callback.message.edit_text(
        "✅ Сохраняю КП в базу данных...\n\n"
        "📅 Укажите сроки выполнения:\n"
        "(Например: '14 дней', '2 недели', '01.02.2024')"
    )
    
    print("[DEBUG] Переход к состоянию waiting_execution_terms")
    
    # Переходим к состоянию ожидания сроков
    await state.set_state(KPStates.waiting_execution_terms)


@router.message(KPStates.waiting_execution_terms)
async def receive_execution_terms(message: Message, state: FSMContext):
    """
    Обработчик ввода сроков выполнения.
    Сохраняет КП в базу данных plita.db.
    """
    execution_terms = message.text.strip()
    print(f"[DEBUG] Получены сроки: {execution_terms}")
    
    # Получаем данные КП из состояния
    data = await state.get_data()
    order_data = data.get('kp_order_data', [])
    xlsx_path = data.get('kp_xlsx_path')
    customer_name = data.get('kp_customer_name')
    manager_name = data.get('kp_manager_name')
    discount_percent = data.get('kp_discount_percent', 0)
    delivery_conditions = data.get('kp_delivery_conditions')
    payment_conditions = data.get('kp_payment_conditions')
    offer_date = data.get('kp_offer_date')
    
    print(f"[DEBUG] Данные из state:")
    print(f"  - Клиент: {customer_name}")
    print(f"  - Менеджер: {manager_name}")
    print(f"  - Скидка: {discount_percent}%")
    print(f"  - Плит в заказе: {len(order_data)}")
    print(f"  - XLSX путь: {xlsx_path}")
    
    try:
        # Сохраняем КП в базу данных
        kp_id = kp_db.save_kp_to_db(
            creation_date=offer_date,
            order_data=order_data,
            xlsx_file_path=xlsx_path,
            customer_name=customer_name,
            manager_name=manager_name,
            discount_percent=discount_percent,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            execution_terms=execution_terms,
            status='в работе'
        )
        
        # Вычисляем общую сумму для отображения
        subtotal = 0.0
        for item in order_data:
            qty = item.get('qty', 0)
            unit_price = item.get('unit_price', 0.0)
            discounted_price = unit_price * (1 - discount_percent / 100)
            subtotal += discounted_price * qty
        
        vat_amount = round(subtotal * 0.20, 2)
        total_amount = round(subtotal + vat_amount, 2)
        
        await message.answer(
            f"✅ КП успешно сохранено в базу данных!\n\n"
            f"📋 Информация о КП:\n"
            f"  • Номер КП: {kp_id}\n"
            f"  • Дата: {offer_date}\n"
            f"  • Клиент: {customer_name}\n"
            f"  • Менеджер: {manager_name}\n"
            f"  • Сумма: {total_amount:,.2f} ₽ (с НДС)\n"
            f"  • Сроки: {execution_terms}\n"
            f"  • Статус: в работе\n\n"
            f"💡 Вы можете отслеживать статус этого КП в базе данных.",
            reply_markup=main_menu_kb()
        )
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при сохранении КП в БД: {str(e)}\n\n"
            "Попробуйте снова позже.",
            reply_markup=main_menu_kb()
        )
        import traceback
        traceback.print_exc()
    
    finally:
        # Очищаем состояние
        await state.clear()


@router.callback_query(F.data == "skip_save_kp")
async def callback_skip_save_kp(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "Не сохранять".
    Пропускает сохранение КП в БД.
    """
    # Убираем "часики" с кнопки
    await callback.answer()
    
    # Редактируем сообщение с кнопками
    await callback.message.edit_text(
        "❌ КП не сохранено в базу данных."
    )
    
    await callback.message.answer(
        "✅ Работа с КП завершена!",
        reply_markup=main_menu_kb()
    )
    
    # Очищаем состояние
    await state.clear()

