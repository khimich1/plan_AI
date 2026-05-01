"""Экспорт диаграммы Ганта и сохранение планов"""
import asyncio
import logging
import os
from datetime import datetime, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import CallbackQuery, FSInputFile
from aiogram.fsm.context import FSMContext

# Импорты из твоего проекта
import sys
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.gantt_excel import create_gantt_excel
from core import kp_db
from core.plan_commit import PlanCommitError, commit_plan_plates

from ..keyboards import main_menu_kb, calendar_days_kb, production_menu_kb
from ..bot_config import OUTPUTS_DIR_STR

# Импорт менеджера планов
from .plan_manager import (
    get_active_plan_id, add_tracks_to_plan, format_plan_stats_message,
    get_all_tracks_from_plan, get_global_days_info, get_global_day_occupancy,
    MAX_TRACKS_PER_DAY, get_all_plans_gantt_data, convert_lookup_keys_to_tuples,
    save_plan, update_plan_metadata, set_active_plan, get_plan_path,
    count_day_tracks
)

router = Router()


@router.callback_query(F.data == "export_gantt_current")
async def export_gantt_current_plan(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "📈 Диаграмма этого плана".
    Строит диаграмму Ганта по ТЕКУЩЕМУ плану в памяти (FSM state), даже если план ещё не сохранён.
    """
    await callback.message.answer("📈 Создаю диаграмму Ганта по текущему плану...")

    data = await state.get_data()

    all_tracks_list = data.get('all_tracks_list') or []
    tracks_count = data.get('tracks_count')
    plate_lookup_exact = data.get('plate_lookup_exact') or {}
    plate_lookup_by_length = data.get('plate_lookup_by_length') or {}
    plan_start_date = data.get('plan_start_date')

    if not all_tracks_list or not tracks_count:
        await callback.message.answer(
            "❌ Текущий план не найден в памяти.\n\n"
            "💡 Сначала выполните планирование:\n"
            "1️⃣ Планирование производства → «🚀 Начать планирование»\n"
            "2️⃣ Выберите КП и дождитесь сообщения «✅ План готов!»\n"
            "3️⃣ Затем нажмите «📈 Диаграмма этого плана»"
        )
        await callback.answer()
        return

    # Парсим дату начала (нужна для дат в Excel)
    start_date_for_gantt = datetime.now()
    if plan_start_date:
        try:
            start_date_for_gantt = datetime.strptime(str(plan_start_date)[:10], '%Y-%m-%d')
        except Exception:
            pass

    try:
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_list,
            tracks_count=int(tracks_count),
            plate_lookup_exact=convert_lookup_keys_to_tuples(plate_lookup_exact),
            plate_lookup_by_length=convert_lookup_keys_to_tuples(plate_lookup_by_length),
            output_dir=OUTPUTS_DIR_STR,
            start_date=start_date_for_gantt
        )

        if gantt_path and os.path.exists(gantt_path):
            total_days = data.get('total_days')
            total_tracks_count = data.get('total_tracks_count') or len(all_tracks_list)

            start_str = start_date_for_gantt.strftime('%d.%m.%Y')
            end_str = ''
            if isinstance(total_days, int) and total_days > 0:
                end_dt = start_date_for_gantt + timedelta(days=total_days - 1)
                end_str = end_dt.strftime('%d.%m.%Y')

            caption_lines = [
                "📈 Диаграмма Ганта этого плана (ещё не сохранённого)\n",
                f"📅 Дата начала: {start_str}",
            ]
            if end_str:
                caption_lines.append(f"📅 Период: {start_str} — {end_str}")
            if total_days:
                caption_lines.append(f"📆 Дней: {total_days}")
            caption_lines.append(f"🛤️ Дорожек: {total_tracks_count}")
            caption_lines.append("\nПодсказка:")
            caption_lines.append("• «📊 Диаграмма Ганта» — это суммарно по ВСЕМ сохранённым планам")
            caption = "\n".join(caption_lines)

            await callback.message.answer_document(
                FSInputFile(gantt_path),
                caption=caption
            )
        else:
            await callback.message.answer(
                "⚠️ Не удалось создать диаграмму.\n"
                "Возможно, в текущем плане нет данных для построения."
            )

    except Exception as e:
        logger.exception(f"Ошибка создания диаграммы текущего плана: {e}")
        await callback.message.answer(
            "❌ Не удалось создать диаграмму Ганта по текущему плану.\n"
            "Подробности в logs/bot.log."
        )

    # Возвращаем клавиатуру выбора дней (чтобы не теряться)
    total_days_state = data.get('total_days', 0)
    plan_start_date_state = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = data.get('days_info', {})
    from_saved_plan = data.get('from_saved_plan', False)

    if isinstance(total_days_state, int) and total_days_state > 0:
        await callback.message.answer(
            "Выберите день для просмотра:",
            reply_markup=calendar_days_kb(
                total_days_state,
                plan_start_date_state,
                completed_days,
                days_info,
                show_save_button=not from_saved_plan
            )
        )

    await callback.answer()


@router.callback_query(F.data == "export_gantt")
async def export_gantt_chart(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "Диаграмма Ганта".
    Создаёт СУММАРНУЮ Excel-диаграмму по ВСЕМ сохранённым планам производства.
    
    Простыми словами:
    - Загружает ВСЕ сохранённые планы
    - Собирает все дорожки и информацию о КП
    - Создаёт одну большую диаграмму Ганта по всем планам
    """
    await callback.message.answer("📊 Создаю суммарную диаграмму Ганта по всем планам...")
    
    # Получаем данные из ВСЕХ планов
    gantt_data = get_all_plans_gantt_data()
    
    if not gantt_data:
        await callback.message.answer(
            "❌ Нет сохранённых планов для создания диаграммы.\n\n"
            "💡 Сначала создайте и сохраните план:\n"
            "1️⃣ Нажмите «🚀 Начать планирование»\n"
            "2️⃣ Выберите КП для производства\n"
            "3️⃣ Нажмите «💾 Сохранить план»",
            reply_markup=main_menu_kb()
        )
        await callback.answer()
        return
    
    # Извлекаем данные
    all_tracks_list = gantt_data['all_tracks']
    plate_lookup_exact = gantt_data['plate_lookup_exact']
    plate_lookup_by_length = gantt_data['plate_lookup_by_length']
    start_date_for_gantt = gantt_data['earliest_start_date']
    plans_count = gantt_data['plans_count']
    total_days = gantt_data['total_days']
    
    # Для корректного подсчёта дней используем среднее количество дорожек
    # (это нужно для совместимости с create_gantt_excel)
    tracks_count = 3  # Среднее значение, не критично для диаграммы
    
    try:
        # Создаём диаграмму Ганта
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_list,
            tracks_count=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            output_dir=OUTPUTS_DIR_STR,
            start_date=start_date_for_gantt
        )
        
        if gantt_path and os.path.exists(gantt_path):
            # Форматируем даты для отображения
            start_date_str = start_date_for_gantt.strftime('%d.%m.%Y')
            end_date_str = gantt_data['latest_end_date'].strftime('%d.%m.%Y')
            
            # Считаем количество уникальных КП в диаграмме
            # (это можно сделать только после создания файла, но для простоты опустим)
            
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
                )
            )
            
            logger.info(f"[GANTT] Диаграмма успешно создана: {gantt_path}")
        else:
            await callback.message.answer(
                "⚠️ Не удалось создать диаграмму.\n"
                "Возможно, нет данных о КП в сохранённых планах.\n\n"
                "💡 Убедитесь, что планы содержат информацию о заказах."
            )
    
    except Exception as e:
        logger.exception(f"Ошибка создания диаграммы: {e}")
        await callback.message.answer(
            "❌ Не удалось создать диаграмму Ганта.\n"
            "Подробности в logs/bot.log."
        )
    
    # Получаем данные из state для возврата к календарю
    data = await state.get_data()
    total_days_state = data.get('total_days', 1)
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    completed_days = data.get('completed_days', [])
    days_info = data.get('days_info', {})
    from_saved_plan = data.get('from_saved_plan', False)
    
    # Показываем клавиатуру выбора дней снова (с датами)
    # только если мы в контексте просмотра календаря
    if total_days_state and total_days_state > 0:
        await callback.message.answer(
            "Выберите день для просмотра:",
            reply_markup=calendar_days_kb(
                total_days_state, 
                plan_start_date, 
                completed_days,
                days_info,
                show_save_button=not from_saved_plan
            )
        )
    
    await callback.answer()


@router.callback_query(F.data == "save_current_plan")
async def save_current_plan(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик сохранения актуального плана производства.
    
    НОВАЯ ЛОГИКА:
    - Добавляет дорожки к активному плану (не перезаписывает!)
    - Если активного плана нет — создаёт новый
    - Показывает статистику: какие дни обновлены, какие созданы
    - Генерирует диаграмму Ганта
    - БЛОКИРУЕТ сохранение при превышении лимита дорожек
    """
    # Получаем данные из state
    data = await state.get_data()
    all_tracks_list = data.get('all_tracks_list', [])
    tracks_count = data.get('tracks_count', 1)
    plate_lookup_exact = data.get('plate_lookup_exact', {})
    plate_lookup_by_length = data.get('plate_lookup_by_length', {})
    
    # Дата начала плана
    plan_start_date = data.get('plan_start_date', datetime.now().strftime('%Y-%m-%d'))
    
    # Дополнительные данные для плана
    orders_2d = data.get('orders_2d', [])
    optimization_result = data.get('optimization_result', {})

    if not all_tracks_list:
        await callback.message.answer(
            "❌ Нет данных для сохранения.\n"
            "Сначала выполните анализ производства.",
            reply_markup=main_menu_kb()
        )
        await callback.answer()
        return
    
    # === ПРОВЕРКА ПРЕВЫШЕНИЯ ЛИМИТА ===
    # ИСПРАВЛЕНИЕ: Берём ID плана ТОЛЬКО из state, без fallback на глобальный!
    # Это гарантирует создание нового плана при active_plan_id=None
    active_plan_id = data.get('active_plan_id')  # Может быть None — это нормально!
    
    # Исключаем текущий план из подсчёта занятости (чтобы не считать дважды)
    # Это исправляет баг, когда при просмотре существующего плана и попытке сохранения
    # система считала дорожки текущего плана дважды
    global_occupancy = get_global_day_occupancy(exclude_plan_id=active_plan_id)
    
    total_days = data.get('total_days', 1)
    try:
        start_dt = datetime.strptime(plan_start_date, '%Y-%m-%d')
    except:
        start_dt = datetime.now()
    
    overloaded_days = []
    for day_num in range(1, total_days + 1):
        day_date = start_dt + timedelta(days=day_num - 1)
        date_key = day_date.strftime('%Y-%m-%d')
        date_display = day_date.strftime('%d.%m')
        
        current_occupied = global_occupancy.get(date_key, 0)
        free_slots = MAX_TRACKS_PER_DAY - current_occupied
        
        if tracks_count > free_slots:
            overloaded_days.append({
                'date': date_display,
                'occupied': current_occupied,
                'free': free_slots,
                'want': tracks_count
            })
    
    # Если есть превышение - БЛОКИРУЕМ сохранение
    if overloaded_days:
        error_lines = ["❌ НЕЛЬЗЯ СОХРАНИТЬ ПЛАН!\n"]
        error_lines.append(f"Превышен лимит дорожек ({MAX_TRACKS_PER_DAY}/день):\n")
        
        for day in overloaded_days[:5]:
            error_lines.append(
                f"  • {day['date']}: занято {day['occupied']}/5, "
                f"свободно {day['free']}, нужно {day['want']}"
            )
        
        if len(overloaded_days) > 5:
            error_lines.append(f"  ... и ещё {len(overloaded_days) - 5} дней")
        
        error_lines.append(f"\n💡 Что делать:")
        error_lines.append(f"1️⃣ Начните планирование заново с меньшим кол-вом дорожек")
        error_lines.append(f"2️⃣ Или выберите другую дату начала")
        error_lines.append(f"3️⃣ Или удалите/отредактируйте другие планы")
        
        await callback.message.answer('\n'.join(error_lines))
        await callback.answer("⚠️ Превышен лимит!")
        return
    
    await callback.message.answer("💾 Сохраняю дорожки в план...")
    
    plan_saved = False  # Флаг для отката
    plan_id = None
    
    try:
        
        # ШАБЛОН ИСПРАВЛЕНИЯ: Подготавливаем план БЕЗ сохранения на диск
        updated_plan, stats = add_tracks_to_plan(
            plan_id=active_plan_id,
            new_tracks_list=all_tracks_list,
            start_date=plan_start_date,
            tracks_per_day=tracks_count,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            orders_2d=orders_2d,
            optimization_result=optimization_result,
            auto_save=False  # НЕ сохраняем автоматически!
        )
        
        db_path = str(PROJECT_ROOT / "plita.db")
        plan_id = updated_plan['id']

        # P9: собираем tracks_by_day из готового плана и передаём в commit,
        # чтобы у каждой плиты в kp_plates был day_number и у каждого
        # track item — kp_plate_id. Без этого срабатывает legacy-ветка
        # (day_number=NULL) и плиты «зависают» вне day_view / complete_day.
        tracks_by_day_for_commit: dict[str, list[dict]] = {}
        for date_key, day_data in (updated_plan.get('days') or {}).items():
            day_number = int((day_data or {}).get('day_number') or 0)
            day_tracks = (day_data or {}).get('tracks') or []
            for track in day_tracks:
                if isinstance(track, dict):
                    track.setdefault('production_day', day_number)
            tracks_by_day_for_commit[date_key] = day_tracks

        try:
            commit_result = commit_plan_plates(
                plan_id=plan_id,
                orders_2d=orders_2d,
                optimization_result=optimization_result,
                all_tracks_list=all_tracks_list,
                db_path=db_path,
                tracks_by_day=tracks_by_day_for_commit,
            )
        except PlanCommitError as commit_error:
            logger.error("[SAVE_PLAN] Не удалось закоммитить план: %s", commit_error)
            raise

        logger.info(
            "[SAVE_PLAN] Помечено %s плит как 'в плане' для плана %s (пропущено %s)",
            commit_result.plates_marked,
            plan_id,
            commit_result.plates_skipped,
        )

        save_plan(updated_plan)
        update_plan_metadata(updated_plan)
        set_active_plan(plan_id)
        plan_saved = True  # Отмечаем, что план сохранён
        logger.info(f"[SAVE_PLAN] План {plan_id} успешно сохранён на диск")
        
        # Сохраняем ID активного плана в state и устанавливаем флаг from_saved_plan
        await state.update_data(
            active_plan_id=updated_plan['id'],
            from_saved_plan=True  # План теперь сохранён
        )
        
        # Парсим дату начала для диаграммы Ганта
        start_date_for_gantt = datetime.now()
        if plan_start_date:
            try:
                start_date_for_gantt = datetime.strptime(plan_start_date, '%Y-%m-%d')
            except ValueError:
                pass
        
        # Собираем все дорожки из обновлённого плана для диаграммы Ганта
        all_tracks_for_gantt = get_all_tracks_from_plan(updated_plan)
        
        # Создаём диаграмму Ганта
        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=all_tracks_for_gantt,
            tracks_count=tracks_count,
            plate_lookup_exact=convert_lookup_keys_to_tuples(updated_plan.get('plate_lookup_exact', {})),
            plate_lookup_by_length=convert_lookup_keys_to_tuples(updated_plan.get('plate_lookup_by_length', {})),
            output_dir=OUTPUTS_DIR_STR,
            start_date=start_date_for_gantt
        )
        
        # Формируем сообщение со статистикой
        stats_message = format_plan_stats_message(stats)
        
        # Подсчитываем общую статистику плана
        total_days = len(updated_plan.get('days', {}))
        total_tracks = sum(
            count_day_tracks(day)
            for day in updated_plan.get('days', {}).values()
        )
        
        # Форматируем дату начала для отображения
        start_date_display = plan_start_date
        try:
            start_date_display = datetime.strptime(plan_start_date, '%Y-%m-%d').strftime('%d.%m.%Y')
        except:
            pass
        
        # Сообщение об успешном сохранении
        success_message = (
            f"✅ План успешно сохранён!\n\n"
            f"{stats_message}\n\n"
            f"📋 План: {updated_plan.get('name', 'Без названия')}\n"
            f"📅 Дата начала: {start_date_display}\n"
            f"📊 Всего дней: {total_days}\n"
            f"🛤️ Всего дорожек: {total_tracks}\n\n"
            f"⭐ Этот план установлен как АКТИВНЫЙ\n"
            f"При входе в «Календарный план» из меню откроется именно он.\n\n"
            f"💡 Как открыть:\n"
            f"Планирование производства → Календарный план"
        )
        
        await callback.message.answer(success_message)
        
        # Выходим в меню планирования производства
        await state.clear()
        await callback.message.answer(
            "📋 Планирование производства плит\n\n"
            "Выберите действие:",
            reply_markup=production_menu_kb()
        )
        
    except Exception as e:
        logger.exception(f"Ошибка при сохранении плана: {e}")
        
        # === ОТКАТ ИЗМЕНЕНИЙ ===
        # Если плиты были помечены, но сохранение плана не удалось - возвращаем плиты
        if plan_id:
            try:
                db_path = str(PROJECT_ROOT / "plita.db")
                recovered = kp_db.return_plan_plates_to_production(plan_id, db_path)
                if recovered > 0:
                    logger.info(f"[ROLLBACK] Возвращено {recovered} плит в производство")
            except Exception as rollback_error:
                logger.error(f"[ROLLBACK] Ошибка при откате плит: {rollback_error}")
        
        # Если план был сохранён на диск, но потом произошла ошибка - удаляем файл
        if plan_saved and plan_id:
            try:
                plan_path = get_plan_path(plan_id)
                if plan_path.exists():
                    os.remove(plan_path)
                    logger.info(f"[ROLLBACK] Удалён файл плана {plan_id}")
            except Exception as delete_error:
                logger.error(f"[ROLLBACK] Ошибка при удалении файла плана: {delete_error}")
        
        await callback.message.answer(
            "❌ Не удалось сохранить план.\n"
            "Все изменения отменены.\n"
            "Подробности в logs/bot.log."
        )
    
    await callback.answer()
