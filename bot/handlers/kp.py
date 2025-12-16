"""Обработчики получения КП (текст + фото с OCR)"""
import asyncio
import os
import sys
import math
from pathlib import Path
from collections import Counter, defaultdict

from aiogram import Router, F
from aiogram.types import Message, FSInputFile
from aiogram.fsm.context import FSMContext

# Добавляем корень проекта в sys.path
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.visualization import visualize_plan
from core.config_and_data import set_plate_lists_from_text
import core.config_and_data as cfg
from core.ocr_gpt import recognize_text_smart, GPT_AVAILABLE, EASYOCR_AVAILABLE
from core.reinforcement_db import get_reinforcement

from ..keyboards import main_menu_kb
from ..states import KPStates
from ..bot_config import OUTPUTS_DIR_STR

router = Router()

# Глобальный кэш для результатов оптимизации
optimization_results_by_load = {}


@router.message(F.text == "Получить КП")
async def btn_get_kp(message: Message, state: FSMContext):
    """Кнопка 'Получить КП' - запрашивает список плит"""
    await state.set_state(KPStates.waiting_for_plate_list)
    
    # Формируем подсказку о фото в зависимости от доступных методов OCR
    if GPT_AVAILABLE:
        photo_hint = "\n📸 Или отправьте фото таблицы - я распознаю через 🧠 GPT-4o (точность 95%+)!"
    elif EASYOCR_AVAILABLE:
        photo_hint = "\n📸 Или отправьте фото таблицы - я распознаю через 🤖 EasyOCR!"
    else:
        photo_hint = ""
    
    await message.answer(
        "✍️ Пришлите список плит в свободной форме.\n"
        "Например: '1.2×3.39 — 2 шт; 0.32×6.63 — 4 шт; 0.32×7.83 — 3 шт'\n"
        f"{photo_hint}\n\n"
        "Я выполню расчёт с оптимизацией и пришлю схемы и смету.\n"
        "💡 Используется каскадная оптимизация для экономии материала!",
        reply_markup=main_menu_kb()
    )


@router.message(KPStates.waiting_for_plate_list, F.photo)
async def receive_photo_with_plates(message: Message, state: FSMContext):
    """
    🧠 УМНАЯ обработка фотографий с плитами:
    1. Скачивает фото
    2. Пробует бесплатный EasyOCR
    3. Если не получилось — использует платный GPT-4o
    4. Парсит распознанный текст
    5. Обрабатывает заказ
    """
    # Проверяем доступность хотя бы одного метода OCR
    if not EASYOCR_AVAILABLE and not GPT_AVAILABLE:
        await message.answer(
            "❌ OCR недоступен. Установите одну из библиотек:\n\n"
            "🤖 EasyOCR (бесплатно):\n"
            "   pip install easyocr\n\n"
            "🧠 GPT-4o (платно, но точнее):\n"
            "   pip install openai\n"
            "   Добавьте в .env: OPENAI_API_KEY=sk-...\n\n"
            "Или отправьте текст заказа вручную."
        )
        return
    
    # Скачиваем фото (берём самое большое разрешение)
    photo = message.photo[-1]
    user_id = message.from_user.id
    os.makedirs("tmp", exist_ok=True)
    photo_path = os.path.join("tmp", f"{user_id}_photo.jpg")
    
    await message.answer("📸 Получил фото! Анализирую...")
    
    try:
        # Скачиваем фото
        await message.bot.download(photo, destination=photo_path)
        
        # 🔥 УМНОЕ РАСПОЗНАВАНИЕ (EasyOCR → GPT fallback)
        result = await recognize_text_smart(photo_path, show_cost=True)
        
        if not result:
            await message.answer(
                "❌ Не удалось распознать текст на фото.\n\n"
                "💡 Попробуйте:\n"
                "• Сделать фото при хорошем освещении\n"
                "• Убедиться, что текст чёткий и читаемый\n"
                "• Расположить камеру параллельно таблице\n"
                "• Отправить текст заказа вручную"
            )
            # Удаляем временный файл
            if os.path.exists(photo_path):
                try:
                    os.remove(photo_path)
                except:
                    pass
            return
        
        # Показываем пользователю результат
        method_emoji = {"EasyOCR": "🤖", "GPT-4o": "🧠"}
        emoji = method_emoji.get(result['method'], "🔍")
        
        # Формируем красивое сообщение
        confidence_percent = int(result['confidence'] * 100)
        status_msg = f"{emoji} **{result['method']}** (уверенность {confidence_percent}%)\n\n"
        
        # Добавляем инфо о стоимости, если использовали GPT
        if result['cost_usd'] > 0:
            rub_cost = result['cost_usd'] * 75
            status_msg += f"💰 Стоимость: ${result['cost_usd']:.4f} (~{rub_cost:.2f}₽)\n\n"
        
        status_msg += f"📋 Распознанный текст:\n```\n{result['text']}\n```\n\nПродолжаю обработку..."
        
        await message.answer(status_msg, parse_mode="Markdown")
        
        # Используем распознанный текст
        cleaned_text = result['text']
        
        # Обрабатываем текст как обычный заказ
        await process_plate_order(message, state, cleaned_text)
        
    except Exception as e:
        await message.answer(f"❌ Ошибка при обработке фото: {str(e)}\n\nПопробуйте отправить текст заказа вручную.")
    finally:
        # Удаляем временный файл
        if os.path.exists(photo_path):
            try:
                os.remove(photo_path)
            except:
                pass


@router.message(KPStates.waiting_for_plate_list)
async def receive_plate_list_and_build(message: Message, state: FSMContext):
    """Обработчик текстового списка плит"""
    user_text = message.text or ""
    await process_plate_order(message, state, user_text)


async def process_plate_order(message: Message, state: FSMContext, user_text: str):
    """
    Общая функция обработки заказа (текст или распознанный с фото)
    """
    await message.answer("⏳ Считаю КП по вашему списку... Это может занять время.")
    
    try:
        # 1) Парсим список пользователя в структуры визуализатора
        unparsed_lines = set_plate_lists_from_text(user_text)
        
        # Если какие‑то строки не распознаны по формату — сразу честно говорим об этом
        if unparsed_lines:
            warn_text = "⚠️ Некоторые строки я не смог распознать по формату и пропустил:\n"
            warn_text += "\n".join(f"• {line}" for line in unparsed_lines[:5])  # Показываем первые 5
            if len(unparsed_lines) > 5:
                warn_text += f"\n... и ещё {len(unparsed_lines) - 5} строк"
            warn_text += (
                "\n\nЯ понимаю, например, такие форматы:\n"
                "• 1.2×3.39 — 2 шт\n"
                "• 0,32x6,63 - 4\n"
                "• Плиты ПБ 78-12-8п 3\n"
                "• ПБ 66,2-12-8п 6\n"
            )
            await message.answer(warn_text)
        
        # ✅ НОВАЯ ЛОГИКА: Группируем плиты по АРМИРОВАНИЮ (из БД)
        orders_by_reinforcement = defaultdict(list)  # {reinforcement_value: [orders_2d]}
        
        # Путь к БД
        db_path = Path(__file__).parent.parent / "pb.db"
        
        print(f"[BOT] Проверяем PLATE_LOAD_DETAILS: {len(cfg.PLATE_LOAD_DETAILS)} записей")
        
        # Используем детальную карту с нагрузками (если есть)
        if cfg.PLATE_LOAD_DETAILS:
            print("[BOT] ✅ Используем PLATE_LOAD_DETAILS (с нагрузками)")
            for (length, width_m, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
                width_mm = int(round(width_m * 1000))
                
                # 🔥 ПОЛУЧАЕМ АРМИРОВАНИЕ ИЗ БД по (длина, нагрузка)
                reinforcement_value = get_reinforcement(
                    length_m=length,
                    load_code=load_code,
                    source='series',  # по серии
                    db_path=db_path,
                    allow_fallback=True
                )
                
                # Если не нашли в БД - используем fallback (группируем по нагрузке как раньше)
                if reinforcement_value is None:
                    reinforcement_key = f"load_{math.floor(load_code)}"
                    print(f"  ⚠️ {qty}x {length}м × {width_mm}мм, нагрузка {cfg.format_reinforcement_from_load_code(load_code)} → армирование НЕ НАЙДЕНО (fallback к группировке по нагрузке)")
                else:
                    # Округляем до 1 знака: 5.23 → 5.2 (чтобы близкие значения попали в одну группу)
                    reinforcement_key = round(reinforcement_value, 1)
                    print(f"  ✅ {qty}x {length}м × {width_mm}мм, нагрузка {cfg.format_reinforcement_from_load_code(load_code)} → армирование {reinforcement_key}")
                
                orders_by_reinforcement[reinforcement_key].append({
                    'length': length,
                    'width': width_mm,
                    'qty': qty,
                    'load_code': load_code,  # Сохраняем нагрузку для отображения
                    'reinforcement': reinforcement_value  # Сохраняем армирование
                })
        else:
            # Fallback: Если PLATE_LOAD_DETAILS пуст
            print("[BOT] ⚠️ PLATE_LOAD_DETAILS пуст, используем fallback (все плиты = 8п)")
            for width_mm, plates_list, target_name in [
                (1200, cfg.PLATES_1_2, 'PLATES_1_2'), (1080, cfg.PLATES_1_08, 'PLATES_1_08'), (1000, cfg.PLATES_1_0, 'PLATES_1_0'),
                (320, cfg.PLATES_0_32, 'PLATES_0_32'), (460, cfg.PLATES_0_46, 'PLATES_0_46'), (700, cfg.PLATES_0_70, 'PLATES_0_70'),
                (720, cfg.PLATES_0_72, 'PLATES_0_72'), (860, cfg.PLATES_0_86, 'PLATES_0_86'), (880, cfg.PLATES_0_88, 'PLATES_0_88'),
                (740, cfg.PLATES_0_74, 'PLATES_0_74'), (480, cfg.PLATES_0_48, 'PLATES_0_48'), (500, cfg.PLATES_0_50, 'PLATES_0_50'),
                (340, cfg.PLATES_0_34, 'PLATES_0_34')
            ]:
                if plates_list:
                    length_counts = Counter(plates_list)
                    for length, qty in length_counts.items():
                        exact_width_m = cfg.get_exact_width(length, target_name, width_mm / 1000.0)
                        exact_width_mm = int(round(exact_width_m * 1000))
                        load_code = cfg.get_load_code_for_plate(length, exact_width_m, default=8)
                        
                        # Получаем армирование из БД
                        reinforcement_value = get_reinforcement(
                            length_m=length,
                            load_code=load_code,
                            source='series',
                            db_path=db_path,
                            allow_fallback=True
                        )
                        
                        if reinforcement_value is None:
                            reinforcement_key = f"load_{load_code}"
                        else:
                            reinforcement_key = round(reinforcement_value, 1)
                        
                        orders_by_reinforcement[reinforcement_key].append({
                            'length': length,
                            'width': exact_width_mm,
                            'qty': qty,
                            'load_code': load_code,
                            'reinforcement': reinforcement_value
                        })
        
        # Если после парсинга не осталось ни одной плиты — сразу выходим
        if not orders_by_reinforcement:
            await message.answer(
                "❌ Не удалось распознать ни одной плиты в вашем сообщении.\n"
                "Проверьте формат строк (ширина×длина×кол-во или 'Плиты ПБ 78-12-8п 3')."
            )
            await state.clear()
            return
        
        # ✅ ЗАПУСКАЕМ ОПТИМИЗАЦИЮ ДЛЯ КАЖДОГО АРМИРОВАНИЯ ОТДЕЛЬНО
        # Безопасная сортировка: сначала числа, потом строки (чтобы не ломалось при смешанных типах)
        keys_list = list(orders_by_reinforcement.keys())
        numeric_keys = sorted([k for k in keys_list if isinstance(k, (int, float))])
        string_keys = sorted([k for k in keys_list if isinstance(k, str)])
        all_keys_sorted = numeric_keys + string_keys
        print(f"\n[BOT] Найдено {len(orders_by_reinforcement)} групп(ы) по армированию: {all_keys_sorted}")
        
        optimization_results_by_reinforcement = {}
        total_plates_all = 0
        total_cost_all = 0
        
        # Создаём карту армирование→оригинальные нагрузки (для правильного отображения)
        reinforcement_to_loads = {}
        for reinforcement_key, orders in orders_by_reinforcement.items():
            loads = set(o['load_code'] for o in orders)
            reinforcement_to_loads[reinforcement_key] = sorted(loads)
        
        # Используем безопасную отсортированную версию ключей
        for reinforcement_key in all_keys_sorted:
            orders_2d = orders_by_reinforcement[reinforcement_key]
            
            # Для отображения собираем все нагрузки в этой группе
            loads_in_group = reinforcement_to_loads[reinforcement_key]
            load_display_list = [cfg.format_reinforcement_from_load_code(lc) for lc in loads_in_group]
            load_display = ", ".join(load_display_list) if len(load_display_list) > 1 else load_display_list[0]
            
            # Красивое отображение ключа группы
            if isinstance(reinforcement_key, (int, float)):
                group_label = f"армирование {reinforcement_key}"
            else:
                group_label = f"{reinforcement_key}"
            
            print(f"\n[BOT] === Оптимизация для {group_label} (нагрузки: {load_display}) ===")
            print(f"[BOT] Плит: {sum(o['qty'] for o in orders_2d)} шт, типов: {len(orders_2d)}")
            
            try:
                from core.optimization import optimize_with_cascading_longitudinal_cuts
                optimization_result = await asyncio.to_thread(
                    optimize_with_cascading_longitudinal_cuts,
                    orders_2d=orders_2d
                )
                
                if optimization_result and optimization_result.get('total_plates', 0) > 0:
                    # Сохраняем с информацией о группе
                    optimization_result['reinforcement_key'] = reinforcement_key
                    optimization_result['loads_in_group'] = loads_in_group
                    optimization_results_by_reinforcement[reinforcement_key] = optimization_result
                    total_plates_all += optimization_result.get('total_plates', 0)
                    total_cost_all += optimization_result.get('total_cost', 0)
                    
                    print(f"[BOT] ✅ {group_label} ({load_display}): {optimization_result['total_plates']} плит, "
                          f"{optimization_result.get('total_cost', 0):,} ₽".replace(',', ' '))
            except Exception as e:
                print(f"[BOT] ❌ Ошибка оптимизации для {group_label}: {e}")
        
        # Сохраняем результаты в глобальную переменную
        if optimization_results_by_reinforcement:
            import core.optimization as optimization
            optimization.OPT_CASCADING_PLAN_BY_LOAD = optimization_results_by_reinforcement
            print(f"\n[BOT] ✅ Сохранено {len(optimization_results_by_reinforcement)} результатов оптимизации")
            
            # ✅ Создаём маппинг нагрузка → армирование для быстрого поиска плана
            load_to_reinforcement_map = {}
            for reinforcement_key, result in optimization_results_by_reinforcement.items():
                loads_in_group = result.get('loads_in_group', [])
                for load_code in loads_in_group:
                    # Если одна нагрузка встречается в нескольких группах - сохраняем список
                    if load_code not in load_to_reinforcement_map:
                        load_to_reinforcement_map[load_code] = []
                    load_to_reinforcement_map[load_code].append(reinforcement_key)
            
            optimization.LOAD_TO_REINFORCEMENT_MAP = load_to_reinforcement_map
            print(f"[BOT] ✅ Создан маппинг нагрузок → армирование: {load_to_reinforcement_map}")
            
            # Показываем сводку пользователю
            opt_msg = "💡 **Результат оптимизации по армированию:**\n"
            # Используем безопасную сортировку результатов
            result_keys = list(optimization_results_by_reinforcement.keys())
            result_numeric = sorted([k for k in result_keys if isinstance(k, (int, float))])
            result_strings = sorted([k for k in result_keys if isinstance(k, str)])
            
            total_track_length = 0.0  # Общая длина всех дорожек
            
            for reinforcement_key in result_numeric + result_strings:
                result = optimization_results_by_reinforcement[reinforcement_key]
                loads_in_group = result.get('loads_in_group', [])
                load_display_list = [cfg.format_reinforcement_from_load_code(lc) for lc in loads_in_group]
                load_display = ", ".join(load_display_list)
                
                if isinstance(reinforcement_key, (int, float)):
                    label = f"арм. {reinforcement_key}"
                else:
                    label = str(reinforcement_key)
                
                # Рассчитываем длину дорожки (сумма длин ИСХОДНЫХ плит, без двойного счёта резов)
                track_length = 0.0
                
                # Приоритет 1: Берём из primary_cuts (первичные резы из исходных плит)
                if result.get('primary_cuts'):
                    for prim_cut in result['primary_cuts']:
                        lengths = prim_cut.get('lengths', [])
                        qty = prim_cut.get('qty', 0)
                        
                        if lengths:
                            # Средняя длина плит в этом резе
                            avg_length = sum(lengths) / len(lengths)
                            track_length += avg_length * qty
                
                # Приоритет 2: Fallback на orders_requested (если нет primary_cuts)
                elif result.get('orders_requested'):
                    for order in result['orders_requested']:
                        track_length += order.get('length', 0) * order.get('qty', 0)
                
                total_track_length += track_length
                
                # Добавляем длину дорожки к сообщению
                opt_msg += f"• **{label}** ({load_display}): {result['total_plates']} плит, **{track_length:.1f}м**\n"
            
            opt_msg += f"\n**Итого:** {total_plates_all} плит, **{total_track_length:.1f}м**, {total_cost_all:,} ₽\n".replace(',', ' ')
            
            await message.answer(opt_msg, parse_mode="Markdown")
        else:
            print("[BOT] ⚠️ Оптимизация не дала результатов, используем fallback")
        
        # 4) Строим приоритет ширин (запасной вариант, если каскадная не сработала)
        if not optimization_results_by_reinforcement:
            from core.optimization import apply_width_optimization
            apply_width_optimization()
        
        # 5) Запускаем расчёт и визуализацию
        result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR_STR)
        if isinstance(result_paths, tuple) and len(result_paths) >= 2:
            png_path, pdf_path = result_paths
            
            # Извлекаем timestamp из имени PNG
            base = os.path.basename(png_path)
            if 'КЗ_' in base:
                timestamp = base.split('КЗ_', 1)[-1].replace('.png', '')
            else:
                timestamp = base.rsplit('_', 1)[-1].replace('.png', '')
            
            # Возможные имена доп.файлов (поддерживаем оба варианта из визуализатора)
            candidates = [
                os.path.join(OUTPUTS_DIR_STR, f'Ведомость_Дорожка_1_{timestamp}.xlsx'),
                os.path.join(OUTPUTS_DIR_STR, f'Смета_Дорожка_1_{timestamp}.xlsx'),
                os.path.join(OUTPUTS_DIR_STR, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx'),
                os.path.join(OUTPUTS_DIR_STR, f'Ведомость_Дорожка_1_{timestamp}.csv'),
                os.path.join(OUTPUTS_DIR_STR, f'Раскладка_Дорожка_1_{timestamp}.csv'),
            ]
            
            await message.answer("✅ Готово! Отправляю файлы:")
            
            if os.path.exists(png_path):
                await message.answer_document(FSInputFile(png_path))
            if os.path.exists(pdf_path):
                await message.answer_document(FSInputFile(pdf_path))
            
            # Отправляем Excel файлы в правильном порядке
            files_sent = 0
            for p in candidates:
                if os.path.exists(p):
                    await message.answer_document(FSInputFile(p))
                    files_sent += 1
            
            # Формируем итоговое сообщение
            final_msg = "📋 **Итоги:**\n• Схема раскладки готова\n• Ведомость и смета сформированы"
            if optimization_results_by_reinforcement:
                final_msg += "\n\n✨ **Использована оптимизация с каскадными резами**\n• Минимум плит\n• Остатки используются повторно"
            await message.answer(final_msg, parse_mode="Markdown")
        else:
            await message.answer("❌ Ошибка при расчёте КП")
    except Exception as e:
        await message.answer(f"❌ Ошибка: {str(e)}", parse_mode=None)
    finally:
        await state.clear()


from aiogram.filters import Command

@router.message(Command("build_plan"))
async def cmd_build_plan(message: Message):
    """Обработчик команды /build_plan"""
    await message.answer("⏳ Выполняю расчёт дорожки, подожди немного...")
    
    try:
        # Запускаем расчёт в отдельном потоке
        result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR_STR)
        
        if isinstance(result_paths, tuple) and len(result_paths) >= 2:
            png_path, pdf_path = result_paths
            
            # Ищем дополнительные файлы
            # Извлекаем timestamp (всё после "КЗ_")
            base = os.path.basename(png_path)
            if 'КЗ_' in base:
                timestamp = base.split('КЗ_', 1)[-1].replace('.png', '')
            else:
                # Fallback: последняя часть после последнего подчеркивания
                timestamp = base.rsplit('_', 1)[-1].replace('.png', '')
            
            csv_path = os.path.join(OUTPUTS_DIR_STR, f'Раскладка_Дорожка_1_{timestamp}.csv')
            xlsx_path = os.path.join(OUTPUTS_DIR_STR, f'Ведомость_Дорожка_1_{timestamp}.xlsx')
            breakdown_path = os.path.join(OUTPUTS_DIR_STR, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx')
            xlsx_smeta_path = os.path.join(OUTPUTS_DIR_STR, f'Смета_Дорожка_1_{timestamp}.xlsx')
            
            await message.answer("✅ Готово! Отправляю файлы:")
            
            # Отправляем изображение как документ, чтобы избежать PHOTO_INVALID_DIMENSIONS
            if os.path.exists(png_path):
                await message.answer_document(FSInputFile(png_path))
            
            # Отправляем документы
            if os.path.exists(pdf_path):
                await message.answer_document(FSInputFile(pdf_path))
            
            if os.path.exists(xlsx_path):
                await message.answer_document(FSInputFile(xlsx_path))
            
            if os.path.exists(xlsx_smeta_path):
                await message.answer_document(FSInputFile(xlsx_smeta_path))
            
            if os.path.exists(breakdown_path):
                await message.answer_document(FSInputFile(breakdown_path))
            
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
        await message.answer(f"❌ Ошибка: {str(e)}", parse_mode=None)

