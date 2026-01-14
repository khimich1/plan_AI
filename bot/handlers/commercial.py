"""Обработчики создания коммерческих предложений PDF/XLSX"""
import asyncio
import os
import re
import sys
import math
from pathlib import Path
from datetime import datetime
from typing import Dict, Any
from collections import defaultdict

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
from core.reinforcement_db import get_reinforcement
from core.visualization import visualize_plan
from core import kp_db

from ..keyboards import main_menu_kb, conditions_choice_kb, save_to_db_kb, cancel_process_kb, managers_selection_kb
from ..states import KPStates
from ..bot_config import OUTPUTS_DIR_STR

router = Router()

# Кэш заказов пользователей
ORDER_CACHE: Dict[int, list] = {}


@router.message(F.text == "📝 Создать КП")
@router.message(Command("commercial_offer"))
async def btn_commercial_offer(message: Message, state: FSMContext):
    """Обработчик запроса на создание коммерческого предложения - ПОШАГОВЫЙ ОПРОС"""
    # Шаг 1: Запрашиваем выбор менеджера из списка
    # Получаем список менеджеров из БД
    from core.kp_db import get_all_managers
    managers = get_all_managers()
    
    if not managers:
        await message.answer(
            "⚠️ В базе данных нет менеджеров.\n"
            "Обратитесь к администратору.",
            reply_markup=main_menu_kb()
        )
        return
    
    await state.set_state(KPStates.waiting_manager_selection)
    await message.answer(
        "📄 Создание коммерческого предложения\n\n"
        "Шаг 1 из 5: Выберите менеджера",
        reply_markup=managers_selection_kb(managers)
    )


# === ПОШАГОВЫЙ ОПРОС ДЛЯ КОММЕРЧЕСКОГО ПРЕДЛОЖЕНИЯ ===

@router.callback_query(F.data.startswith("select_manager_"))
async def select_manager_callback(callback: CallbackQuery, state: FSMContext):
    """Обработчик выбора менеджера из списка"""
    manager_id = int(callback.data.split("_")[-1])
    
    # Получаем данные менеджера из БД
    from core.kp_db import get_manager_by_id
    manager = get_manager_by_id(manager_id)
    
    if not manager:
        await callback.answer("❌ Менеджер не найден", show_alert=True)
        return
    
    # Сохраняем ВСЕ данные менеджера в состояние
    await state.update_data(
        manager_id=manager['id'],
        manager_name=manager['fio'],
        manager_phone=manager['contact_number'],
        manager_email=manager['email']
    )
    
    await callback.message.edit_text(
        f"✅ Менеджер: {manager['fio']}\n\n"
        "Шаг 2 из 5: Введите имя клиента\n"
        "(Для кого создается коммерческое предложение)"
    )
    
    await state.set_state(KPStates.waiting_client_name)
    await callback.message.answer(
        "Введите имя клиента:",
        reply_markup=main_menu_kb()
    )
    await callback.message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
    )
    await callback.answer()


@router.message(KPStates.waiting_client_name)
async def receive_client_name(message: Message, state: FSMContext):
    """Шаг 2: Получаем имя клиента"""
    client_name = message.text.strip()
    
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
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
    )


@router.message(KPStates.waiting_plates_list)
async def receive_plates_list(message: Message, state: FSMContext):
    """Шаг 3: Получаем список плит (текст или фото)"""
    
    # === ПРОВЕРЯЕМ, ЧТО ПРИШЛО: ТЕКСТ ИЛИ ФОТО ===
    
    if message.photo:
        # 📸 ПРИШЛО ФОТО — используем GPT сразу
        await message.answer("📸 Фото получено! Распознаю через GPT-4o...")
        
        # Скачиваем фото (берём самое большое разрешение)
        photo = message.photo[-1]
        user_id = message.from_user.id
        os.makedirs("tmp", exist_ok=True)
        photo_path = os.path.join("tmp", f"{user_id}_commercial_photo.jpg")
        
        try:
            # Скачиваем фото
            await message.bot.download(photo, destination=photo_path)
            
            # 🔥 РАСПОЗНАВАНИЕ ЧЕРЕЗ GPT (force_gpt=True пропускает EasyOCR)
            from core.ocr_gpt import recognize_text_smart
            result = await recognize_text_smart(photo_path, force_gpt=True, show_cost=True)
            
            if result and result.get('text'):
                plates_text = result['text']
                cost = result.get('cost_usd', 0)
                
                # Показываем результат распознавания
                preview = plates_text[:200] + ('...' if len(plates_text) > 200 else '')
                info_msg = f"✅ Распознано через GPT-4o Vision\n\n📋 Найденные плиты:\n{preview}"
                
                if cost > 0:
                    rub_cost = cost * 75  # Примерный курс
                    info_msg += f"\n\n💰 Стоимость: ${cost:.4f} (~{rub_cost:.2f}₽)"
                
                await message.answer(info_msg)
            else:
                await message.answer(
                    "❌ Не удалось распознать текст на фото.\n"
                    "Попробуйте:\n"
                    "• Сделать фото более чётким\n"
                    "• Прислать текстом\n"
                    "• Использовать формат 'ПБ XX-XX-Xп количество'"
                )
                return
                
        except Exception as e:
            print(f"[COMMERCIAL] Ошибка распознавания фото: {e}")
            import traceback
            traceback.print_exc()
            await message.answer(
                f"❌ Ошибка при обработке фото: {str(e)}\n\n"
                "Попробуйте прислать список текстом."
            )
            return
    
    elif message.text:
        # 📝 ПРИШЁЛ ТЕКСТ — используем как раньше
        plates_text = message.text.strip()
    
    else:
        # ❌ НЕ ТЕКСТ И НЕ ФОТО
        await message.answer(
            "❌ Пришлите список плит текстом или фото таблицы.\n\n"
            "Примеры форматов текста:\n"
            "• '1.2×3.39 — 2 шт'\n"
            "• 'ПБ 38-12-8п 2'"
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return
    
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
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
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
            await message.answer(
                "Или нажмите кнопку ниже для отмены:",
                reply_markup=cancel_process_kb()
            )
            return
    except ValueError:
        await message.answer(
            "❌ Неверный формат числа. Введите просто число (например: 0, 5, 10):",
            reply_markup=main_menu_kb()
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
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
        await callback.message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
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
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
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
    manager_phone = data.get('manager_phone', '')
    manager_email = data.get('manager_email', '')
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
        
        # ✅ ЗАПУСКАЕМ ОПТИМИЗАЦИЮ (упрощённая версия, без разделения по армированию)
        await message.answer("🔄 Оптимизирую раскрой плит для минимизации стоимости...")
        
        # Собираем все плиты в один список orders_2d (без группировки по армированию)
        orders_2d = []
        
        if cfg.PLATE_LOAD_DETAILS:
            print("[COMMERCIAL] ✅ Используем PLATE_LOAD_DETAILS (с нагрузками)")
            for (length, width_m, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
                width_mm = int(round(width_m * 1000))
                
                orders_2d.append({
                    'length': length,
                    'width': width_mm,
                    'qty': qty,
                    'load_code': load_code
                })
        
        if orders_2d:
            print(f"[COMMERCIAL] Всего плит для оптимизации: {sum(o['qty'] for o in orders_2d)} шт, типов: {len(orders_2d)}")
            
            try:
                from core.optimization import optimize_with_cascading_longitudinal_cuts
                import core.optimization as optimization
                
                # Запускаем оптимизацию для ВСЕХ плит сразу (без разделения)
                optimization_result = await asyncio.to_thread(
                    optimize_with_cascading_longitudinal_cuts,
                    orders_2d=orders_2d
                )
                
                if optimization_result and optimization_result.get('total_plates', 0) > 0:
                    # Сохраняем результат в ОБЩИЙ план (не по нагрузкам)
                    optimization.OPT_CASCADING_PLAN = optimization_result
                    
                    # Также сохраняем в BY_LOAD под общим ключом для совместимости
                    # Собираем все нагрузки из заказа
                    all_loads = set(o['load_code'] for o in orders_2d)
                    optimization_result['loads_in_group'] = sorted(all_loads)
                    
                    # Используем специальный ключ 'all' для обозначения, что это общий план
                    optimization.OPT_CASCADING_PLAN_BY_LOAD = {'all': optimization_result}
                    
                    # Создаём маппинг: все нагрузки указывают на общий план
                    optimization.LOAD_TO_REINFORCEMENT_MAP = {
                        load_code: ['all'] for load_code in all_loads
                    }
                    
                    total_plates = optimization_result.get('total_plates', 0)
                    total_cost = optimization_result.get('total_cost', 0)
                    
                    print(f"[COMMERCIAL] ✅ Оптимизация завершена: {total_plates} плит, {total_cost:,} ₽".replace(',', ' '))
                    await message.answer(f"✅ Оптимизация завершена! Использовано {total_plates} исходных плит")
                    
            except Exception as e:
                print(f"[COMMERCIAL] ❌ Ошибка оптимизации: {e}")
                import traceback
                traceback.print_exc()
                # Продолжаем без оптимизации (цены будут посчитаны по старой логике)
        
        # Используем build_price_rows для получения правильных цен
        from viz_modules.procurement import build_price_rows, build_component_breakdown
        from viz_modules.price_utils import load_price_table_from_xlsx
        
        # Загружаем таблицу цен
        price_table = load_price_table_from_xlsx(cfg.PRICE_XLSX_PATH)
        
        # Получаем строки сметы
        price_rows, total_sum = await asyncio.to_thread(
            build_price_rows,
            price_table,
            reinforcement_code=8
        )
        
        # Получаем детальную разбивку компонентов (для PDF)
        breakdown_tables = await asyncio.to_thread(
            build_component_breakdown,
            price_table,
            price_rows
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
        
        # 🆕 ПОЛУЧАЕМ СЛЕДУЮЩИЙ НОМЕР КП ИЗ БД
        from core.kp_db import get_next_kp_number
        kp_db_id = get_next_kp_number()
        print(f"[DEBUG] Предполагаемый номер КП из БД: {kp_db_id}")
        
        # Генерируем PDF с номером КП из БД
        pdf_buffer = await asyncio.to_thread(
            generate_commercial_offer_pdf,
            order_data,
            offer_number,
            offer_date,
            client_name,      # Используем имя клиента из опроса
            manager_name,     # имя менеджера
            manager_phone,    # телефон менеджера
            manager_email,    # email менеджера
            discount_percent, # процент скидки
            kp_db_id          # 🆕 НОМЕР КП ИЗ БД!
        )
        
        # Генерируем XLSX с номером КП из БД
        try:
            xlsx_buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data,
                offer_number,
                offer_date,
                client_name,         # Используем имя клиента
                manager_name,        # имя менеджера
                manager_phone,       # телефон менеджера
                manager_email,       # email менеджера
                discount_percent,    # передаем процент скидки
                delivery_conditions, # условия поставки
                payment_conditions,  # условия оплаты
                kp_db_id             # 🆕 НОМЕР КП ИЗ БД!
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
        
        # Сохраняем детальную разбивку в отдельный Excel файл
        breakdown_filename = f"Детальная_разбивка_{offer_number}_{offer_date.replace('.', '')}.xlsx"
        breakdown_path = os.path.join(OUTPUTS_DIR_STR, breakdown_filename)
        
        with open(pdf_path, 'wb') as f:
            f.write(pdf_buffer.getvalue())
        
        if has_xlsx:
            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_buffer.getvalue())
        
        # Сохраняем детальную разбивку
        has_breakdown = False
        if breakdown_tables:
            from core.commercial_offer import save_breakdown_to_excel
            has_breakdown = await asyncio.to_thread(
                save_breakdown_to_excel,
                breakdown_tables,
                breakdown_path
            )
        
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
        
        # Отправляем детальную разбивку
        if has_breakdown and os.path.exists(breakdown_path):
            await message.answer_document(
                FSInputFile(breakdown_path),
                caption=f"📋 Детальная разбивка компонентов № {offer_number}"
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
            "• НДС (22%)\n"
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
        
        # ✅ ЗАПУСКАЕМ ОПТИМИЗАЦИЮ
        await message.answer("🔄 Оптимизирую раскрой плит для минимизации стоимости...")
        
        # Группируем плиты по армированию (из БД)
        orders_by_reinforcement = defaultdict(list)
        db_path = Path(__file__).parent.parent / "pb.db"
        
        if cfg.PLATE_LOAD_DETAILS:
            print("[COMMERCIAL] ✅ Используем PLATE_LOAD_DETAILS (с нагрузками)")
            for (length, width_m, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
                width_mm = int(round(width_m * 1000))
                
                # Получаем армирование из БД по (длина, нагрузка)
                reinforcement_value = get_reinforcement(
                    length_m=length,
                    load_code=load_code,
                    source='series',
                    db_path=db_path,
                    allow_fallback=True
                )
                
                # Если не нашли в БД - используем fallback (группируем по нагрузке)
                if reinforcement_value is None:
                    reinforcement_key = f"load_{math.floor(load_code)}"
                else:
                    reinforcement_key = round(reinforcement_value, 1)
                
                orders_by_reinforcement[reinforcement_key].append({
                    'length': length,
                    'width': width_mm,
                    'qty': qty,
                    'load_code': load_code,
                    'reinforcement': reinforcement_value
                })
        
        # Запускаем оптимизацию для каждой группы армирования
        optimization_results_by_reinforcement = {}
        
        if orders_by_reinforcement:
            # Безопасная сортировка ключей
            keys_list = list(orders_by_reinforcement.keys())
            numeric_keys = sorted([k for k in keys_list if isinstance(k, (int, float))])
            string_keys = sorted([k for k in keys_list if isinstance(k, str)])
            all_keys_sorted = numeric_keys + string_keys
            
            print(f"[COMMERCIAL] Найдено {len(orders_by_reinforcement)} групп(ы) по армированию")
            
            for reinforcement_key in all_keys_sorted:
                orders_2d = orders_by_reinforcement[reinforcement_key]
                
                try:
                    from core.optimization import optimize_with_cascading_longitudinal_cuts
                    optimization_result = await asyncio.to_thread(
                        optimize_with_cascading_longitudinal_cuts,
                        orders_2d=orders_2d
                    )
                    
                    if optimization_result and optimization_result.get('total_plates', 0) > 0:
                        # Сохраняем с информацией о группе
                        optimization_result['reinforcement_key'] = reinforcement_key
                        loads_in_group = set(o['load_code'] for o in orders_2d)
                        optimization_result['loads_in_group'] = sorted(loads_in_group)
                        optimization_results_by_reinforcement[reinforcement_key] = optimization_result
                        
                        print(f"[COMMERCIAL] ✅ Армирование {reinforcement_key}: {optimization_result['total_plates']} плит")
                except Exception as e:
                    print(f"[COMMERCIAL] ❌ Ошибка оптимизации для армирования {reinforcement_key}: {e}")
            
            # Сохраняем результаты в глобальную переменную
            if optimization_results_by_reinforcement:
                import core.optimization as optimization
                optimization.OPT_CASCADING_PLAN_BY_LOAD = optimization_results_by_reinforcement
                
                # Создаём маппинг нагрузка → армирование для быстрого поиска плана
                load_to_reinforcement_map = {}
                for reinforcement_key, result in optimization_results_by_reinforcement.items():
                    loads_in_group = result.get('loads_in_group', [])
                    for load_code in loads_in_group:
                        if load_code not in load_to_reinforcement_map:
                            load_to_reinforcement_map[load_code] = []
                        load_to_reinforcement_map[load_code].append(reinforcement_key)
                
                optimization.LOAD_TO_REINFORCEMENT_MAP = load_to_reinforcement_map
                print(f"[COMMERCIAL] ✅ Сохранено {len(optimization_results_by_reinforcement)} результатов оптимизации")
                
                await message.answer("✅ Оптимизация завершена! Формирую документы...")
        
        # 🔥 ТЕПЕРЬ build_price_rows получит ОПТИМИЗИРОВАННЫЕ данные из OPT_CASCADING_PLAN_BY_LOAD!
        from viz_modules.procurement import build_price_rows, build_component_breakdown
        from viz_modules.price_utils import load_price_table_from_xlsx
        
        # Загружаем таблицу цен для расчётов
        price_table = load_price_table_from_xlsx(cfg.PRICE_XLSX_PATH)
        
        # Получаем строки сметы с ПРАВИЛЬНЫМИ ценами (С УЧЁТОМ ОПТИМИЗАЦИИ!)
        price_rows, total_sum = await asyncio.to_thread(
            build_price_rows,
            price_table,
            reinforcement_code=8
        )
        
        # Получаем детальную разбивку компонентов (для PDF)
        breakdown_tables = await asyncio.to_thread(
            build_component_breakdown,
            price_table,
            price_rows
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
            customer_name,
            None,  # manager_name - не используется в старом потоке
            None,  # manager_phone
            None,  # manager_email
            0,     # discount_percent - в старом потоке нет скидки
            None   # kp_db_id - старый поток без БД
        )
        
        # Генерируем XLSX
        try:
            xlsx_buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data,
                offer_number,
                offer_date,
                customer_name,
                None,  # manager_name - не используется в старом потоке
                None,  # manager_phone
                None,  # manager_email
                0,     # discount_percent
                None,  # delivery_conditions
                None,  # payment_conditions
                None   # kp_db_id - старый поток без БД
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
        
        # Сохраняем детальную разбивку в отдельный Excel файл
        breakdown_filename = f"Детальная_разбивка_{offer_number}_{offer_date.replace('.', '')}.xlsx"
        breakdown_path = os.path.join(OUTPUTS_DIR_STR, breakdown_filename)
        
        with open(pdf_path, 'wb') as f:
            f.write(pdf_buffer.getvalue())
        
        if has_xlsx:
            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_buffer.getvalue())
        
        # Сохраняем детальную разбивку
        has_breakdown = False
        if breakdown_tables:
            from core.commercial_offer import save_breakdown_to_excel
            has_breakdown = await asyncio.to_thread(
                save_breakdown_to_excel,
                breakdown_tables,
                breakdown_path
            )
        
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
        
        # Отправляем детальную разбивку
        if has_breakdown and os.path.exists(breakdown_path):
            await message.answer_document(
                FSInputFile(breakdown_path),
                caption=f"📋 Детальная разбивка компонентов № {offer_number}"
            )
        
        # 🔥 ГЕНЕРИРУЕМ СХЕМУ И ДЕТАЛЬНУЮ РАЗБИВКУ
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
            # Сохраняем данные КП для последующего сохранения в БД
            await state.update_data(
                kp_order_data=order_data,
                kp_xlsx_path=xlsx_path if has_xlsx and os.path.exists(xlsx_path) else None,
                kp_customer_name=customer_name,
                kp_manager_name=None,  # В старом потоке нет менеджера
                kp_discount_percent=0,  # В старом потоке нет скидки
                kp_delivery_conditions=None,
                kp_payment_conditions=None,
                kp_offer_date=offer_date  # 🔥 ИСПРАВЛЕНИЕ: Сохраняем дату!
            )
            
            # Очищаем состояние FSM, но оставляем данные
            await state.set_state(None)
            
            await message.answer(
                "✨ Документы содержат:\n"
                "• Подробную спецификацию\n"
                "• Расчёт стоимости материалов\n"
                "• Стоимость резов\n"
                "• Вес изделий\n"
                "• НДС (22%)\n"
                "• Условия оплаты\n\n"
                "📊 XLSX файл содержит расчётные формулы Excel!\n"
                "📐 Схема раскладки и детальная разбивка также готовы!",
                reply_markup=main_menu_kb()
            )
            
            # Предлагаем сохранить КП в базу данных
            await message.answer(
                "💾 Хотите сохранить это КП в базу данных?\n\n"
                "Если сохраните, вы сможете отслеживать статус выполнения заказа.",
                reply_markup=save_to_db_kb()
            )
        else:
            await message.answer(
                "❌ Ошибка при сохранении файла",
                reply_markup=main_menu_kb()
            )
            await state.clear()
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при генерации КП: {str(e)}\n\n"
            "Проверьте формат данных и попробуйте снова.",
            reply_markup=main_menu_kb()
        )
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
    execution_terms_input = message.text.strip()
    print(f"[DEBUG] Получены сроки: {execution_terms_input}")
    
    # === ПАРСИМ СРОКИ И ВЫЧИСЛЯЕМ ДАТУ ДЕДЛАЙНА ===
    from datetime import timedelta
    import re
    
    deadline_date = None
    
    # ИСПРАВЛЕНИЕ: Сначала пробуем распознать ДАТУ (чтобы "01.02.2026" не распознавалось как "1 день")
    # Вариант 1: Формат ДД.ММ.ГГГГ (например: "01.02.2026")
    try:
        deadline_date = datetime.strptime(execution_terms_input, '%d.%m.%Y')
        print(f"[DEBUG] Распознана дата (ДД.ММ.ГГГГ): {deadline_date.strftime('%d.%m.%Y')}")
    except ValueError:
        pass
    
    # Вариант 2: Формат ГГГГ-ММ-ДД (например: "2026-02-01")
    if not deadline_date:
        try:
            deadline_date = datetime.strptime(execution_terms_input, '%Y-%m-%d')
            print(f"[DEBUG] Распознана дата (ГГГГ-ММ-ДД): {deadline_date.strftime('%d.%m.%Y')}")
        except ValueError:
            pass
    
    # Вариант 3: Пользователь ввёл количество дней (например: "14", "14 дней", "30дней")
    if not deadline_date:
        match_days = re.search(r'(\d+)\s*(?:дн|день|дней|day|days)', execution_terms_input, re.IGNORECASE)
        if match_days:
            days = int(match_days.group(1))
            deadline_date = datetime.now() + timedelta(days=days)
            print(f"[DEBUG] Распознано {days} дней, дедлайн: {deadline_date.strftime('%d.%m.%Y')}")
    
    # Вариант 4: Пользователь ввёл количество недель (например: "2 недели", "3week")
    if not deadline_date:
        match_weeks = re.search(r'(\d+)\s*(?:нед|недел|недели|week|weeks)', execution_terms_input, re.IGNORECASE)
        if match_weeks:
            weeks = int(match_weeks.group(1))
            deadline_date = datetime.now() + timedelta(weeks=weeks)
            print(f"[DEBUG] Распознано {weeks} недель, дедлайн: {deadline_date.strftime('%d.%m.%Y')}")
    
    # Если не удалось распознать, используем 14 дней по умолчанию
    if not deadline_date:
        deadline_date = datetime.now() + timedelta(days=14)
        await message.answer(
            f"⚠️ Не удалось распознать формат срока.\n"
            f"Использую значение по умолчанию: 14 дней\n"
            f"Дедлайн: {deadline_date.strftime('%d.%m.%Y')}"
        )
    
    # Форматируем дату для сохранения в БД
    execution_terms = deadline_date.strftime('%d.%m.%Y')
    print(f"[DEBUG] Итоговая дата дедлайна: {execution_terms}")
    
    # Получаем данные КП из состояния
    data = await state.get_data()
    order_data = data.get('kp_order_data', [])
    xlsx_path = data.get('kp_xlsx_path')
    customer_name = data.get('kp_customer_name')
    manager_name = data.get('kp_manager_name')
    manager_phone = data.get('manager_phone', '')
    manager_email = data.get('manager_email', '')
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
    print(f"  - Дата КП: {offer_date}")
    
    # 🔥 ПРОВЕРКА: Если нет обязательных данных - показываем ошибку
    if not offer_date or not order_data:
        await message.answer(
            "❌ Не удалось получить данные КП.\n\n"
            "Попробуйте создать КП заново через кнопку '📝 Создать КП'.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
        return
    
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
            execution_terms=execution_terms,  # Теперь это дата в формате ДД.ММ.ГГГГ
            status='в работе'
        )
        
        # 🔥 ИСПРАВЛЕНИЕ: Используем ту же функцию расчета, что и в XLSX и в save_kp_to_db
        # Получаем итоговую сумму из сохраненного КП (она уже рассчитана правильно)
        kp_info = kp_db.get_kp_by_id(kp_id)
        if kp_info:
            total_amount = kp_info.get('total_amount', 0)
        else:
            # Fallback: если не удалось получить из БД, используем ту же функцию расчета
            from core.commercial_offer_xlsx import calculate_total_cost
            totals = calculate_total_cost(order_data, discount_percent)
            total_amount = totals['total_with_vat']
        
        await message.answer(
            f"✅ КП успешно сохранено в базу данных!\n\n"
            f"📋 Информация о КП:\n"
            f"  • Номер КП: {kp_id}\n"
            f"  • Дата: {offer_date}\n"
            f"  • Клиент: {customer_name}\n"
            f"  • Менеджер: {manager_name}\n"
            f"  • Сумма: {total_amount:,.2f} ₽ (с НДС)\n"
            f"  • Срок изготовления до: {execution_terms}\n"
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


@router.callback_query(F.data == "save_kp_to_archive")
async def callback_save_kp_to_archive(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "В архив".
    Сохраняет КП со статусом "в архиве" БЕЗ запроса сроков выполнения.
    
    Простыми словами:
    - Берёт все данные КП из памяти (state)
    - Сохраняет в БД со статусом "в архиве"
    - НЕ запрашивает сроки выполнения (execution_terms = None)
    - КП не попадёт в планирование производства
    """
    print("[DEBUG] Нажата кнопка 'В архив'")
    
    # Убираем "часики" с кнопки
    await callback.answer()
    
    # Редактируем сообщение с кнопками
    await callback.message.edit_text(
        "✅ Сохраняю КП в архив..."
    )
    
    # Получаем данные КП из состояния
    data = await state.get_data()
    order_data = data.get('kp_order_data', [])
    xlsx_path = data.get('kp_xlsx_path')
    customer_name = data.get('kp_customer_name')
    manager_name = data.get('kp_manager_name')
    manager_phone = data.get('manager_phone', '')
    manager_email = data.get('manager_email', '')
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
    print(f"  - Дата КП: {offer_date}")
    
    # Проверка обязательных данных
    if not offer_date or not order_data:
        await callback.message.answer(
            "❌ Не удалось получить данные КП.\n\n"
            "Попробуйте создать КП заново через кнопку '📝 Создать КП'.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
        return
    
    try:
        # Сохраняем КП в базу данных со статусом "в архиве"
        kp_id = kp_db.save_kp_to_db(
            creation_date=offer_date,
            order_data=order_data,
            xlsx_file_path=xlsx_path,
            customer_name=customer_name,
            manager_name=manager_name,
            discount_percent=discount_percent,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            execution_terms=None,  # БЕЗ СРОКОВ!
            status='в архиве'  # СТАТУС АРХИВА!
        )
        
        # Получаем итоговую сумму из сохраненного КП
        kp_info = kp_db.get_kp_by_id(kp_id)
        if kp_info:
            total_amount = kp_info.get('total_amount', 0)
        else:
            # Fallback: рассчитываем вручную
            from core.commercial_offer_xlsx import calculate_total_cost
            totals = calculate_total_cost(order_data, discount_percent)
            total_amount = totals['total_with_vat']
        
        await callback.message.answer(
            f"✅ КП успешно сохранено в архив!\n\n"
            f"📋 Информация о КП:\n"
            f"  • Номер КП: {kp_id}\n"
            f"  • Дата: {offer_date}\n"
            f"  • Клиент: {customer_name}\n"
            f"  • Менеджер: {manager_name}\n"
            f"  • Сумма: {total_amount:,.2f} ₽ (с НДС)\n"
            f"  • Статус: в архиве 📦\n\n"
            f"💡 КП находится в архиве и не попадёт в производство.\n"
            f"Вы можете просмотреть его через кнопку '📁 Архив'.",
            reply_markup=main_menu_kb()
        )
    
    except Exception as e:
        await callback.message.answer(
            f"❌ Ошибка при сохранении КП в архив: {str(e)}\n\n"
            "Попробуйте снова позже.",
            reply_markup=main_menu_kb()
        )
        import traceback
        traceback.print_exc()
    
    finally:
        # Очищаем состояние
        await state.clear()


@router.callback_query(F.data == "cancel_process")
async def cancel_commercial_process(callback: CallbackQuery, state: FSMContext):
    """Отмена процесса создания коммерческого предложения"""
    await state.clear()
    await callback.message.answer(
        "❌ Создание коммерческого предложения отменено.\n"
        "Выберите действие:",
        reply_markup=main_menu_kb()
    )
    await callback.answer("Отменено")

