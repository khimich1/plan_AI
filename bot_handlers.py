import asyncio
import os
from datetime import datetime
from typing import Any, Dict

from aiogram import Router, F
from aiogram.types import (
    Message,
    CallbackQuery,
    FSInputFile,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
)
from aiogram.filters import Command
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext
from visualization import visualize_plan
from config_and_data import set_plate_lists_from_text
from optimization import apply_width_optimization, optimize_with_cascading_longitudinal_cuts
import config_and_data as cfg
from bot_config import OUTPUTS_DIR
from planning import plan_tracks, available_days, track_to_text, render_line
from commercial_offer import generate_commercial_offer_pdf

router = Router()

PLANNING_CACHE: Dict[int, Dict[str, Any]] = {}
ORDER_CACHE: Dict[int, list] = {}  # Кэш для хранения заказов пользователей

def register_handlers(dp):
    """Регистрируем все обработчики"""
    dp.include_router(router)

class KPStates(StatesGroup):
    waiting_for_plate_list = State()
    waiting_for_commercial_offer = State()

def main_menu_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Получить КП"), KeyboardButton(text="Планирование по дням")],
            [KeyboardButton(text="Оптимизация резов")],
            [KeyboardButton(text="Коммерческое предложение PDF")],
        ],
        resize_keyboard=True
    )

@router.message(Command("start"))
async def cmd_start(message: Message):
    """Обработчик команды /start"""
    await message.answer(
        "👋 Привет! Я бот для расчёта и визуализации дорожек ПБ.\n\n"
        "🔧 Что я умею:\n"
        "• Строить планы раскладки плит\n"
        "• Рассчитывать стоимость и отходы\n"
        "• Оптимизировать раскрой (экономия до 40%)\n"
        "• Экспортировать результаты в файлы\n\n"
        "Выберите действие кнопкой ниже или /help для справки",
        reply_markup=main_menu_kb()
    )

@router.message(F.text == "Получить КП")
async def btn_get_kp(message: Message, state: FSMContext):
    await state.set_state(KPStates.waiting_for_plate_list)
    await message.answer(
        "✍️ Пришлите список плит в свободной форме.\n"
        "Например: '1.2×3.39 — 2 шт; 0.32×6.63 — 4 шт; 0.32×7.83 — 3 шт'\n\n"
        "Я выполню расчёт с оптимизацией и пришлю схемы и смету.\n"
        "💡 Используется каскадная оптимизация для экономии материала!",
        reply_markup=main_menu_kb()
    )

@router.message(KPStates.waiting_for_plate_list)
async def receive_plate_list_and_build(message: Message, state: FSMContext):
    # На первом шаге просто принимаем текст как подтверждение и запускаем существующий расчёт
    await message.answer("⏳ Считаю КП по вашему списку... Это может занять время.")
    try:
        # 1) Парсим список пользователя в структуры визуализатора
        set_plate_lists_from_text(message.text or "")
        
        # 2) Собираем заказы для 2D оптимизации (длина + ширина)
        from collections import Counter
        orders_2d = []
        
        # Для каждой ширины группируем плиты по длине
        for width_mm, plates_list in [
            (1200, cfg.PLATES_1_2), (1080, cfg.PLATES_1_08),  # КРИТИЧНО: Плиты БЕЗ реза!
            (320, cfg.PLATES_0_32), (460, cfg.PLATES_0_46), (700, cfg.PLATES_0_70),
            (720, cfg.PLATES_0_72), (860, cfg.PLATES_0_86), (880, cfg.PLATES_0_88),
            (740, cfg.PLATES_0_74), (480, cfg.PLATES_0_48), (500, cfg.PLATES_0_50),
            (340, cfg.PLATES_0_34)
        ]:
            if plates_list:
                # Группируем по длине (плиты с одинаковой длиной объединяем)
                length_counts = Counter(plates_list)
                for length, qty in length_counts.items():
                    orders_2d.append({
                        'length': length,
                        'width': width_mm,
                        'qty': qty
                    })
        
        # Для обратной совместимости сохраняем старый формат (только ширины)
        orders = {}
        for order in orders_2d:
            width = order['width']
            orders[width] = orders.get(width, 0) + order['qty']
        
        # 3) Запускаем 2D оптимизацию (с учётом длины и ширины)
        optimization_result = None
        if orders_2d:
            print(f"[BOT] Запускаем 2D оптимизацию для заказа:")
            for order in orders_2d:
                print(f"  - {order['qty']}x {order['length']}м × {order['width']}мм")
            try:
                from optimization import OPT_CASCADING_PLAN, optimize_with_cascading_longitudinal_cuts
                optimization_result = await asyncio.to_thread(
                    optimize_with_cascading_longitudinal_cuts,
                    orders_2d=orders_2d  # Передаём как именованный параметр для режима 2D
                )
                print(f"[BOT] Получен результат: {optimization_result}")
                if optimization_result and optimization_result.get('total_plates', 0) > 0:
                    # Сохраняем результат в глобальную переменную для визуализации
                    import optimization
                    optimization.OPT_CASCADING_PLAN = optimization_result
                    print(f"[BOT] OK: Результат сохранён в optimization.OPT_CASCADING_PLAN")
                    
                    opt_msg = (
                        "💡 **Результат оптимизации:**\n"
                        f"• Плит потребуется: **{optimization_result['total_plates']} шт**\n"
                        f"• Стоимость: **{optimization_result['total_cost']:,} ₽**\n".replace(',', ' ') +
                        f"• Отходы: **{optimization_result.get('waste_width', 0)} мм**\n"
                    )
                    await message.answer(opt_msg, parse_mode="Markdown")
            except Exception as e:
                # Если оптимизация не сработала, продолжаем со старым методом
                print(f"[Cascading optimization failed]: {e}")
        
        # 4) Строим приоритет ширин (запасной вариант, если каскадная не сработала)
        if not optimization_result:
            apply_width_optimization()
        
        # 5) Запускаем расчёт и визуализацию
        result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR)
        if isinstance(result_paths, tuple) and len(result_paths) >= 2:
            png_path, pdf_path = result_paths

            # Извлекаем timestamp из имени PNG
            base = os.path.basename(png_path)
            # Ожидаемый формат: ..._{timestamp}.png
            timestamp = base.rsplit('_', 1)[-1].replace('.png', '')

            # Возможные имена доп.файлов (поддерживаем оба варианта из визуализатора)
            candidates = [
                os.path.join(OUTPUTS_DIR, f'Ведомость_Дорожка_1_{timestamp}.xlsx'),
                os.path.join(OUTPUTS_DIR, f'Смета_Дорожка_1_{timestamp}.xlsx'),
                os.path.join(OUTPUTS_DIR, f'Ведомость_Дорожка_1_{timestamp}.csv'),
                os.path.join(OUTPUTS_DIR, f'Раскладка_Дорожка_1_{timestamp}.csv'),
            ]

            await message.answer("✅ Готово! Отправляю файлы:")

            if os.path.exists(png_path):
                await message.answer_document(FSInputFile(png_path))
            if os.path.exists(pdf_path):
                await message.answer_document(FSInputFile(pdf_path))
            for p in candidates:
                if os.path.exists(p):
                    await message.answer_document(FSInputFile(p))

            # Формируем итоговое сообщение
            final_msg = "📋 **Итоги:**\n• Схема раскладки готова\n• Ведомость и смета сформированы"
            if optimization_result and optimization_result.get('total_plates', 0) > 0:
                final_msg += "\n\n✨ **Использована оптимизация с каскадными резами**\n• Минимум плит\n• Остатки используются повторно"
            await message.answer(final_msg, parse_mode="Markdown")
        else:
            await message.answer("❌ Ошибка при расчёте КП")
    except Exception as e:
        await message.answer(f"❌ Ошибка: {str(e)}")
    finally:
        await state.clear()

@router.message(Command("build_plan"))
async def cmd_build_plan(message: Message):
    """Обработчик команды /build_plan"""
    await message.answer("⏳ Выполняю расчёт дорожки, подожди немного...")
    
    try:
        # Запускаем расчёт в отдельном потоке
        result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR)
        
        if isinstance(result_paths, tuple) and len(result_paths) >= 2:
            png_path, pdf_path = result_paths
            
            # Ищем дополнительные файлы
            timestamp = os.path.basename(png_path).split('_')[-1].replace('.png', '')
            csv_path = os.path.join(OUTPUTS_DIR, f'Раскладка_Дорожка_1_{timestamp}.csv')
            xlsx_path = os.path.join(OUTPUTS_DIR, f'Ведомость_Дорожка_1_{timestamp}.xlsx')
            
            await message.answer("✅ Готово! Отправляю файлы:")
            
            # Отправляем изображение как документ, чтобы избежать PHOTO_INVALID_DIMENSIONS
            if os.path.exists(png_path):
                await message.answer_document(FSInputFile(png_path))
            
            # Отправляем документы
            if os.path.exists(xlsx_path):
                await message.answer_document(FSInputFile(xlsx_path))
            
            if os.path.exists(pdf_path):
                await message.answer_document(FSInputFile(pdf_path))
            
            if os.path.exists(csv_path):
                await message.answer_document(FSInputFile(csv_path))
            
            await message.answer(
                "📋 **Результаты расчёта готовы!**\n\n"
                "• Схема раскладки сохранена\n"
                "• Ведомость материалов готова\n"
                "• Смета стоимости рассчитана\n"
                "• Все файлы экспортированы"
            )
        else:
            await message.answer("❌ Ошибка при расчёте плана")
            
    except Exception as e:
        await message.answer(f"❌ Ошибка: {str(e)}")

@router.message(Command("help"))
async def cmd_help(message: Message):
    """Обработчик команды /help"""
    help_text = """
📖 **Помощь по командам:**

🏗️ **Построить план** - создаёт визуализацию дорожки с расчётом стоимости

**Команды:**
• `/start` - главное меню
• `/build_plan` - построить план дорожки
• `/optimize` - оптимизация раскроя с экономией до 40%
• `/help` - эта справка
• `/stats` - статистика проекта

**Форматы файлов:**
• PNG - схема раскладки
• PDF - техническая документация  
• XLSX - ведомость и смета
• CSV - данные для импорта

💡 **Оптимизация резов:**
Использует каскадные продольные резы для минимизации отходов и экономии материала.
    """
    await message.answer(help_text, parse_mode="Markdown")

@router.message(Command("stats"))
async def cmd_stats(message: Message):
    """Обработчик команды /stats"""
    try:
        # Подсчитываем файлы в папке outputs
        files_count = len([f for f in os.listdir(OUTPUTS_DIR) if f.endswith(('.png', '.pdf', '.xlsx'))])
        
        stats_text = f"""
📊 **Статистика проекта:**

📁 Файлов создано: {files_count}
📂 Папка результатов: `{OUTPUTS_DIR}`

🔧 **Доступные функции:**
• Визуализация раскладки
• Расчёт стоимости материалов
• Экспорт в различные форматы

📈 **Последние результаты:**
• PNG схемы: {len([f for f in os.listdir(OUTPUTS_DIR) if f.endswith('.png')])} шт
• PDF документы: {len([f for f in os.listdir(OUTPUTS_DIR) if f.endswith('.pdf')])} шт
• Excel файлы: {len([f for f in os.listdir(OUTPUTS_DIR) if f.endswith('.xlsx')])} шт
        """
        
        await message.answer(stats_text, parse_mode="Markdown")
        
    except Exception as e:
        await message.answer(f"❌ Ошибка: {str(e)}")

@router.message(Command("optimize"))
@router.message(F.text == "Оптимизация резов")
async def cmd_optimize(message: Message):
    """Оптимизация раскроя с каскадными продольными резами"""
    await message.answer("⏳ Выполняю оптимизацию раскроя с учётом вторичных резов...")
    
    try:
        # Собираем заказы из текущей конфигурации
        orders = {}
        if cfg.PLATES_0_32:
            orders[320] = len(cfg.PLATES_0_32)
        if cfg.PLATES_0_46:
            orders[460] = len(cfg.PLATES_0_46)
        if cfg.PLATES_0_70:
            orders[700] = len(cfg.PLATES_0_70)
        if cfg.PLATES_0_72:
            orders[720] = len(cfg.PLATES_0_72)
        if cfg.PLATES_0_86:
            orders[860] = len(cfg.PLATES_0_86)
        if cfg.PLATES_0_88:
            orders[880] = len(cfg.PLATES_0_88)
        if cfg.PLATES_0_74:
            orders[740] = len(cfg.PLATES_0_74)
        if cfg.PLATES_0_48:
            orders[480] = len(cfg.PLATES_0_48)
        if cfg.PLATES_0_50:
            orders[500] = len(cfg.PLATES_0_50)
        if cfg.PLATES_0_34:
            orders[340] = len(cfg.PLATES_0_34)
        
        if not orders:
            await message.answer(
                "⚠️ Нет данных для оптимизации.\n"
                "Сначала используйте 'Получить КП' для загрузки списка плит.",
                reply_markup=main_menu_kb()
            )
            return
        
        # Запускаем оптимизацию в отдельном потоке
        result = await asyncio.to_thread(optimize_with_cascading_longitudinal_cuts, orders)
        
        if result and result.get('total_plates', 0) > 0:
            # Формируем красивый ответ
            response = "✅ **Оптимизация завершена!**\n\n"
            response += f"📊 **Результат:**\n"
            response += f"• Плит потребуется: **{result['total_plates']} шт**\n"
            response += f"• Стоимость: **{result['total_cost']:,} ₽**\n".replace(',', ' ')
            response += f"• Отходы по ширине: **{result.get('waste_width', 0)} мм**\n\n"
            
            if result.get('primary_cuts'):
                response += "🔹 **Первичные резы:**\n"
                for cut in result['primary_cuts']:
                    response += f"  • {cut['qty']} плит → {cut['width']} мм + остаток {cut['rest']} мм\n"
            
            if result.get('secondary_cuts'):
                response += f"\n🔸 **Вторичные резы (из остатков):**\n"
                for cut in result['secondary_cuts']:
                    if cut.get('pieces', 1) > 1:
                        response += f"  • {cut['qty']} остатков {cut['source']} мм → {cut['pieces']} частей по {cut['cuts'][0]} мм\n"
                    else:
                        cuts_str = ' + '.join(str(c) for c in cut['cuts'])
                        response += f"  • {cut['qty']} остатков {cut['source']} мм → {cuts_str} мм\n"
            
            response += "\n💡 **Преимущества:**\n"
            response += "• Минимум плит\n"
            response += "• Остатки используются повторно\n"
            response += "• Меньше отходов\n"
            
            await message.answer(response, parse_mode="Markdown", reply_markup=main_menu_kb())
        else:
            await message.answer(
                "❌ Не удалось выполнить оптимизацию.\n"
                "Проверьте корректность данных.",
                reply_markup=main_menu_kb()
            )
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при оптимизации: {str(e)}\n\n"
            f"Убедитесь, что библиотека PuLP установлена.",
            reply_markup=main_menu_kb()
        )


@router.message(F.text == "Планирование по дням")
async def btn_planning_days(message: Message):
    await message.answer("⏳ Строю календарь дорожек… подождите пару секунд.")

    try:
        schedule, report_path = await asyncio.to_thread(plan_tracks)

        if not schedule:
            await message.answer(
                "⚠️ Не найдено плит в базе. Попробуйте обновить данные.",
                reply_markup=main_menu_kb(),
            )
            return

        PLANNING_CACHE[message.from_user.id] = {
            "schedule": schedule,
            "report": report_path,
        }

        days = available_days(schedule)
        buttons = [
            [InlineKeyboardButton(text=f"День {day}", callback_data=f"plan_day:{day}")]
            for day in days
        ]

        summary_lines = [
            f"День {day}: {sum(1 for t in schedule if t.day == day)} дорожек"
            for day in days
        ]

        keyboard = InlineKeyboardMarkup(inline_keyboard=buttons)

        await message.answer(
            "✅ План готов!\n\n" + "\n".join(summary_lines) + "\n\nВыберите день:",
            reply_markup=keyboard,
        )

        if report_path and report_path.exists():
            await message.answer_document(FSInputFile(report_path))

    except Exception as e:
        await message.answer(
            f"❌ Ошибка при планировании: {e}",
            reply_markup=main_menu_kb(),
        )


@router.callback_query(F.data.startswith("plan_day:"))
async def cb_plan_day(callback: CallbackQuery):
    await callback.answer()

    cache = PLANNING_CACHE.get(callback.from_user.id)
    if not cache:
        await callback.message.answer(
            "⚠️ План не найден. Нажмите «Планирование по дням» ещё раз.",
            reply_markup=main_menu_kb(),
        )
        return

    try:
        day = int(callback.data.split(":", 1)[1])
    except (ValueError, AttributeError):
        await callback.message.answer("❌ Не удалось определить день.")
        return

    schedule = cache.get("schedule")
    if not schedule:
        await callback.message.answer("⚠️ План пуст. Постройте его заново.")
        return

    day_tracks = [track for track in schedule if track.day == day]
    if not day_tracks:
        await callback.message.answer(f"⚠️ На день {day} дорожек нет.")
        return

    await callback.message.answer(f"📍 День {day}: готовлю визуализации по линиям…")

    for track in sorted(day_tracks, key=lambda t: t.line):
        await callback.message.answer(track_to_text(track), parse_mode="Markdown")
        try:
            png_path, pdf_path, extras = await asyncio.to_thread(render_line, track)
        except Exception as e:
            await callback.message.answer(f"❌ Ошибка визуализации линии {track.line}: {e}")
            continue

        if png_path.exists():
            await callback.message.answer_document(
                FSInputFile(str(png_path)), caption=f"День {day} • Линия {track.line}"
            )
        if pdf_path.exists():
            await callback.message.answer_document(FSInputFile(str(pdf_path)))
        for extra in extras:
            await callback.message.answer_document(FSInputFile(str(extra)))


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
        # Парсим список пользователя
        set_plate_lists_from_text(message.text or "")
        
        # Собираем данные заказа из глобальных списков
        from collections import Counter
        order_data = []
        
        # Собираем все плиты по типам
        plate_groups = [
            (1200, cfg.PLATES_1_2, "12"),
            (1080, cfg.PLATES_1_08, "10.8"),
            (1000, cfg.PLATES_1_0, "10"),
            (320, cfg.PLATES_0_32, "3.2"),
            (460, cfg.PLATES_0_46, "4.6"),
            (700, cfg.PLATES_0_70, "7"),
            (720, cfg.PLATES_0_72, "7.2"),
            (860, cfg.PLATES_0_86, "8.6"),
            (880, cfg.PLATES_0_88, "8.8"),
            (740, cfg.PLATES_0_74, "7.4"),
            (480, cfg.PLATES_0_48, "4.8"),
            (500, cfg.PLATES_0_50, "5"),
            (340, cfg.PLATES_0_34, "3.4"),
        ]
        
        for width_mm, plates_list, width_dm_str in plate_groups:
            if plates_list:
                # Группируем по длине
                length_counts = Counter(plates_list)
                for length_m, qty in length_counts.items():
                    length_dm = int(round(length_m * 10))
                    # Формируем наименование в формате "Плиты ПБ 38-12-8п"
                    if width_mm >= 1000:
                        width_str = str(int(round(width_mm / 100)))
                    else:
                        # Для малых ширин используем дм с точкой
                        width_str = width_dm_str.replace('.', ',')
                    
                    name = f"Плиты ПБ {length_dm}-{width_str}-8п"
                    
                    order_data.append({
                        "name": name,
                        "length_m": length_m,
                        "width_m": width_mm / 1000.0,  # переводим в метры
                        "qty": qty
                    })
        
        if not order_data:
            await message.answer(
                "❌ Не удалось распознать плиты в вашем сообщении.\n"
                "Пожалуйста, проверьте формат и попробуйте снова.",
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
        
        # Генерируем PDF в памяти
        pdf_buffer = await asyncio.to_thread(
            generate_commercial_offer_pdf,
            order_data,
            offer_number,
            offer_date,
            customer_name
        )
        
        # Сохраняем во временный файл для отправки
        pdf_filename = f"КП_{offer_number}_{offer_date.replace('.', '')}.pdf"
        pdf_path = os.path.join(OUTPUTS_DIR, pdf_filename)
        
        with open(pdf_path, 'wb') as f:
            f.write(pdf_buffer.getvalue())
        
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
                caption=f"📄 Коммерческое предложение № {offer_number}"
            )
            await message.answer(
                "✨ Документ содержит:\n"
                "• Подробную спецификацию\n"
                "• Расчёт стоимости материалов\n"
                "• Стоимость резов\n"
                "• Вес изделий\n"
                "• НДС (20%)\n"
                "• Условия оплаты",
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


# ==================== НОВЫЕ КОМАНДЫ: /myorders, /export ====================

@router.message(Command("myorders"))
async def cmd_myorders(message: Message):
    """Показывает историю заказов пользователя"""
    try:
        import sqlite3
        from domain.export import get_user_orders
        
        con = sqlite3.connect('pb.db')
        orders = get_user_orders(con, message.from_user.id, limit=10)
        con.close()
        
        if not orders:
            await message.answer(
                "📋 У вас пока нет сохранённых заказов.\n\n"
                "Создайте заказ через 'Получить КП' или 'Коммерческое предложение PDF'",
                reply_markup=main_menu_kb()
            )
            return
        
        # Формируем список заказов
        response = "📋 <b>История ваших заказов:</b>\n\n"
        
        for order in orders:
            status_icon = {
                'created': '🆕',
                'processing': '⏳',
                'completed': '✅',
                'archived': '📦'
            }.get(order['status'], '❓')
            
            client_info = f" ({order['client_name']})" if order['client_name'] else ""
            
            response += (
                f"{status_icon} <b>Заказ #{order['id']}</b>{client_info}\n"
                f"   Дата: {order['created_at'][:10]}\n"
                f"   Позиций: {order['items_count']}\n"
                f"   /export_{order['id']} - экспортировать\n\n"
            )
        
        response += "\n💡 Для экспорта заказа используйте команду /export_НОМЕР"
        
        await message.answer(response, parse_mode="HTML", reply_markup=main_menu_kb())
        
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при получении истории заказов: {str(e)}",
            reply_markup=main_menu_kb()
        )


@router.message(Command("export"))
async def cmd_export(message: Message):
    """Экспортирует заказ в ZIP архив"""
    try:
        # Парсим ID заказа из команды /export_123
        command_parts = message.text.split('_')
        if len(command_parts) < 2:
            await message.answer(
                "❓ Укажите номер заказа: /export_123\n\n"
                "Посмотреть список заказов: /myorders",
                reply_markup=main_menu_kb()
            )
            return
        
        try:
            order_id = int(command_parts[1])
        except ValueError:
            await message.answer(
                "❌ Неверный формат номера заказа",
                reply_markup=main_menu_kb()
            )
            return
        
        import sqlite3
        from pathlib import Path
        from domain.export import get_order_items, create_order_archive
        from domain.calc import cost_standard, cost_addon
        from domain.excel_kz import generate_kz_excel
        from commercial_offer import generate_commercial_offer_pdf
        from datetime import datetime
        
        # Получаем данные заказа
        con = sqlite3.connect('pb.db')
        items = get_order_items(con, order_id)
        
        if not items:
            con.close()
            await message.answer(
                f"❌ Заказ #{order_id} не найден или у вас нет к нему доступа",
                reply_markup=main_menu_kb()
            )
            return
        
        await message.answer("⏳ Формирую архив заказа...")
        
        # Генерируем файлы
        output_dir = Path("Визуализация_Раскладки")
        output_dir.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 1. Excel КЗ
        excel_path = generate_kz_excel(
            con,
            items,
            tracks=None,
            output_path=str(output_dir / f"kz_{order_id}_{timestamp}.xlsx"),
            order_number=str(order_id),
            customer_name=None
        )
        
        # 2. PDF КП
        order_data = []
        for item in items:
            length_dm = int(round(item['length_m'] * 10))
            width_dm = int(round(item['width_m'] * 10))
            name = f"ПБ {length_dm}-{width_dm}-{int(item['load_class'])}п"
            order_data.append({
                'name': name,
                'length_m': item['length_m'],
                'width_m': item['width_m'],
                'qty': item['qty']
            })
        
        pdf_buffer = generate_commercial_offer_pdf(
            order_data,
            offer_number=str(order_id),
            offer_date=datetime.now().strftime("%d.%m.%Y"),
            customer_name=None
        )
        
        pdf_path = output_dir / f"kp_{order_id}_{timestamp}.pdf"
        with open(pdf_path, 'wb') as f:
            f.write(pdf_buffer.getvalue())
        
        con.close()
        
        # 3. Архивируем
        archive_path = create_order_archive(
            order_id,
            [excel_path, pdf_path],
            output_dir=str(output_dir)
        )
        
        # Отправляем архив
        if archive_path.exists():
            await message.answer_document(
                FSInputFile(archive_path),
                caption=f"📦 Архив заказа #{order_id}\n\nВключает КП (PDF) и КЗ (Excel)"
            )
            
            await message.answer(
                "✅ Архив готов!\n\n"
                "💡 Хотите отправить на email? Напишите адрес в ответ на это сообщение",
                reply_markup=main_menu_kb()
            )
        else:
            await message.answer(
                "❌ Ошибка при создании архива",
                reply_markup=main_menu_kb()
            )
        
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при экспорте: {str(e)}",
            reply_markup=main_menu_kb()
        )
