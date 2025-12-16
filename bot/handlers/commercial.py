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
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext

# Добавляем корень проекта в sys.path
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.config_and_data import set_plate_lists_from_text
import core.config_and_data as cfg
from core.commercial_offer import generate_commercial_offer_pdf, generate_commercial_offer_xlsx
from core.reinforcement_db import get_reinforcement

from ..keyboards import main_menu_kb
from ..states import KPStates
from ..bot_config import OUTPUTS_DIR_STR

router = Router()

# Кэш заказов пользователей
ORDER_CACHE: Dict[int, list] = {}


@router.message(F.text == "Коммерческое предложение PDF")
@router.message(Command("commercial_offer"))
async def btn_commercial_offer(message: Message, state: FSMContext):
    """Обработчик запроса на создание коммерческого предложения"""
    await state.set_state(KPStates.waiting_for_commercial_offer)
    await message.answer(
        "📄 Создание коммерческого предложения\n\n"
        "Пришлите список плит в свободной форме.\n\n"
        "Примеры форматов:\n"
        "• '1.2×3.39 — 2 шт'\n"
        "• '0.32×6.63 — 4 шт'\n"
        "• 'ПБ 38-12-8п 2'\n"
        "• 'ПБ 66-3-8п 4'\n\n"
        "Я создам PDF с расчётом стоимости, веса и НДС.",
        reply_markup=main_menu_kb()
    )


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
        
        # ✅ ЗАПУСКАЕМ ОПТИМИЗАЦИЮ (как в "Получить КП")
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
        from viz_modules.procurement import build_price_rows
        from viz_modules.price_utils import load_price_table_from_xlsx
        
        # Загружаем таблицу цен для расчётов
        price_table = load_price_table_from_xlsx(cfg.PRICE_XLSX_PATH)
        
        # Получаем строки сметы с ПРАВИЛЬНЫМИ ценами (С УЧЁТОМ ОПТИМИЗАЦИИ!)
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
            
        if os.path.exists(pdf_path):
            await message.answer(
                "✨ Документы содержат:\n"
                "• Подробную спецификацию\n"
                "• Расчёт стоимости материалов\n"
                "• Стоимость резов\n"
                "• Вес изделий\n"
                "• НДС (20%)\n"
                "• Условия оплаты\n\n"
                "📊 XLSX файл содержит расчётные формулы Excel!",
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

