"""Основная логика планирования производства — thin adapter над core pipeline."""
import logging
import math
from collections import Counter
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

from aiogram import Router
from aiogram.types import Message
from aiogram.fsm.context import FSMContext

import sys

BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as cfg
from app.services.production_planning_service import (
    ProductionPlanBuildError,
    ProductionPlanningService,
)
from bot.services.production_planning_adapter import (
    apply_rest_matching,
    enrich_lookup_for_secondary_cuts,
    load_kp_list_for_bot_filter,
    load_plates_for_production,
    normalize_kp_plate_ids,
    rebuild_load_result_for_plates,
)
from core.db_config import PLITA_DB_PATH
from core.optimization.result_contract import is_optimization_success
from core.plate_order_context import PlateOrderContext

from ..keyboards import main_menu_kb, calendar_days_kb
from ..states import ProductionStates
from .plan_manager import get_global_day_occupancy, MAX_TRACKS_PER_DAY
from core.work_calendar import nth_working_day

router = Router()
production_planning_service = ProductionPlanningService()


async def load_and_plan_production(
    message: Message,
    state: FSMContext,
    plate_order_ctx: PlateOrderContext,
):
    """
    Универсальная функция загрузки КП и планирования производства.
    Работает с разными способами фильтрации: date, kp, all, customer.
    """
    data = await state.get_data()
    tracks_count = data.get("tracks_count", 1)
    filter_method = data.get("filter_method", "date")
    plan_start_date = data.get("plan_start_date", datetime.now().strftime("%Y-%m-%d"))
    completed_days = data.get("completed_days", [])
    target_date_str = data.get("target_date")

    target_date = None
    if filter_method == "date" and target_date_str:
        target_date = datetime.fromisoformat(target_date_str)

    kp_plate_ids = normalize_kp_plate_ids(data.get("kp_plate_ids"))

    kp_list = load_kp_list_for_bot_filter(
        db_path=str(PLITA_DB_PATH),
        filter_method=filter_method,
        target_date=target_date,
        kp_ids=data.get("kp_ids", []) if filter_method == "kp" else None,
        customer_name=data.get("customer_name", "") if filter_method == "customer" else "",
    )

    if not kp_list:
        await message.answer(
            "❌ Нет подходящих КП для производства.",
            reply_markup=main_menu_kb(),
        )
        await state.clear()
        return

    await message.answer(f"✅ Найдено КП: {len(kp_list)}\nЗагружаю плиты...")

    try:
        try:
            load_result = load_plates_for_production(
                production_planning_service,
                kp_list=kp_list,
                kp_plate_ids=kp_plate_ids,
                start_date=plan_start_date,
                tracks_count=tracks_count,
            )
        except ProductionPlanBuildError as exc:
            await message.answer(
                f"❌ {exc}",
                reply_markup=main_menu_kb(),
            )
            await state.clear()
            return

        plates_from_rests, plates_for_optimizer, all_from_rests = apply_rest_matching(
            production_planning_service,
            load_result.selected_plates,
            db_path=str(PROJECT_ROOT / "plita.db"),
        )

        if all_from_rests:
            await message.answer(
                "✅ Все плиты можно взять из остатков!\n"
                "Оптимизация не требуется.",
                reply_markup=main_menu_kb(),
            )
            await state.update_data(
                plates_from_rests=plates_from_rests,
                all_from_rests=True,
            )
            await state.clear()
            return

        if plates_from_rests:
            total_from_rests = sum(p["qty"] for p in plates_from_rests)
            rests_msg = f"📦 Плиты из остатков: {total_from_rests} шт\n\n"
            for p in plates_from_rests:
                rests_msg += f"✅ {p['plate_name']} × {p['qty']} (КП #{p['kp_id']})\n"
                rests_msg += f"   Из остатка: {p['rest_length']}м × {p['rest_width_mm']}мм\n"
                if p["match_type"] == "exact":
                    rests_msg += "   Точное совпадение, себестоимость: 0 руб.\n"
                elif p["match_type"] == "width_cut":
                    rests_msg += f"   Резы: продольный ({p['length_m']}м)\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                elif p["match_type"] == "length_cut":
                    rests_msg += "   Резы: поперечный\n"
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                else:
                    rests_msg += (
                        f"   Резы: продольный ({p['length_m']}м), поперечный\n"
                    )
                    rests_msg += f"   Себестоимость: {p['cut_cost']:.0f} руб.\n"
                rests_msg += "\n"
            rests_msg += "💰 Эти плиты уже оплачены - чистая прибыль!"
            await message.answer(rests_msg)

        load_result = rebuild_load_result_for_plates(
            load_result, plates_for_optimizer
        )
        orders_2d = load_result.orders_2d

        if not orders_2d:
            await message.answer(
                "❌ Не найдено плит для планирования.",
                reply_markup=main_menu_kb(),
            )
            await state.clear()
            return

        await message.answer("⏳ Запускаю оптимизацию раскроя...")

        load_result, opt_result = production_planning_service.run_planning_pipeline(
            load_result=load_result,
            plate_order_ctx=plate_order_ctx,
        )
        optimization_result = opt_result.optimization_result
        all_tracks_list = opt_result.all_tracks_list

        if (
            not is_optimization_success(optimization_result)
            or optimization_result.get("total_plates", 0) == 0
            or not all_tracks_list
        ):
            await message.answer(
                "❌ Оптимизация не дала результатов.",
                reply_markup=main_menu_kb(),
            )
            await state.clear()
            return

        await message.answer(
            f"✅ Оптимизация завершена! "
            f"Исходных плит: {optimization_result.get('total_plates', 0)}"
        )

        plate_lookup_exact = enrich_lookup_for_secondary_cuts(
            load_result.plate_lookup_exact,
            load_result.plate_lookup_by_length,
            orders_2d,
            optimization_result,
        )
        plate_lookup_by_length = load_result.plate_lookup_by_length

        plate_order_ctx.load_production_snapshot(orders_2d, optimization_result)

        await message.answer("⏳ Подсчитываю дорожки...")

        total_tracks_count = len(all_tracks_list)
        total_days = math.ceil(total_tracks_count / tracks_count)

        plan_plates = _count_plan_plates(all_tracks_list)
        kp_ids_in_production = list(
            {p.get("kp_id") for p in load_result.selected_plates if p.get("kp_id")}
        )
        if kp_ids_in_production and plan_plates:
            _log_plan_vs_db_mismatch(kp_ids_in_production, plan_plates)

        plan_start_display = plan_start_date
        try:
            plan_start_display = datetime.strptime(
                plan_start_date, "%Y-%m-%d"
            ).strftime("%d.%m.%Y")
        except ValueError:
            pass

        await message.answer(
            f"✅ План готов!\n\n"
            f"📊 Параметры:\n"
            f"  • Дата начала: {plan_start_display}\n"
            f"  • Всего дорожек: {total_tracks_count}\n"
            f"  • Дорожек в день: {tracks_count}\n"
            f"  • Потребуется дней: {total_days}\n\n"
            f"💡 Что дальше?\n"
            f"1️⃣ Просмотрите дни ниже 👇\n"
            f"2️⃣ Чтобы посмотреть диаграмму ДО сохранения — нажмите «📈 Диаграмма этого плана»\n"
            f"3️⃣ Нажмите «💾 Сохранить план» когда всё готово\n\n"
            f"⚠️ ВАЖНО: План сохраняется только после нажатия кнопки!\n"
            f"Без сохранения он останется только в памяти.\n\n"
            f"«📈 Диаграмма этого плана» — по текущему расчёту (даже без сохранения)."
        )

        global_occupancy = get_global_day_occupancy()
        days_info: dict[str, dict[str, object]] = {}
        try:
            start_dt = datetime.strptime(plan_start_date, "%Y-%m-%d")
        except ValueError:
            start_dt = datetime.now()

        overloaded_days: list[dict[str, object]] = []
        for day_num in range(1, total_days + 1):
            day_date = datetime.combine(
                nth_working_day(start_dt.date(), day_num),
                datetime.min.time(),
            )
            date_key = day_date.strftime("%Y-%m-%d")
            date_display = day_date.strftime("%d.%m")
            current_occupied = global_occupancy.get(date_key, 0)
            free_slots = MAX_TRACKS_PER_DAY - current_occupied

            if tracks_count > free_slots:
                overloaded_days.append(
                    {
                        "date": date_display,
                        "occupied": current_occupied,
                        "free": free_slots,
                        "want": tracks_count,
                        "excess": tracks_count - free_slots,
                    }
                )

            days_info[date_key] = {
                "occupied": current_occupied,
                "max": MAX_TRACKS_PER_DAY,
                "completed": False,
                "day_number": day_num,
            }

        if overloaded_days:
            warning_lines = ["⚠️ ВНИМАНИЕ! Превышение лимита дорожек!\n"]
            warning_lines.append(
                f"Вы хотите планировать {tracks_count} дорожек/день,"
            )
            warning_lines.append("но на некоторых датах не хватает места:\n")
            for day in overloaded_days[:5]:
                warning_lines.append(
                    f"  • {day['date']}: занято {day['occupied']}/5, "
                    f"свободно {day['free']}, нужно {day['want']}"
                )
            if len(overloaded_days) > 5:
                warning_lines.append(
                    f"  ... и ещё {len(overloaded_days) - 5} дней"
                )
            warning_lines.append("\n💡 Решения:")
            warning_lines.append("1️⃣ Уменьшите количество дорожек в день")
            warning_lines.append("2️⃣ Выберите другую дату начала")
            warning_lines.append("3️⃣ Удалите или отредактируйте другие планы")
            await message.answer("\n".join(warning_lines))

        await state.update_data(
            total_tracks_count=total_tracks_count,
            total_days=total_days,
            tracks_count=tracks_count,
            all_tracks_list=all_tracks_list,
            orders_2d=orders_2d,
            plate_lookup_exact=plate_lookup_exact,
            plate_lookup_by_length=plate_lookup_by_length,
            optimization_result=optimization_result,
            target_date=target_date_str,
            plates_from_rests=plates_from_rests,
            plan_start_date=plan_start_date,
            completed_days=completed_days,
            days_info=days_info,
        )

        await message.answer(
            "Выберите день производства:",
            reply_markup=calendar_days_kb(
                total_days,
                plan_start_date,
                completed_days,
                days_info,
            ),
        )
        await state.set_state(ProductionStates.waiting_day_selection)

    except Exception as e:
        logger.exception("Ошибка при планировании производства: %s", e)
        await message.answer(
            "❌ Ошибка при планировании производства.\n\n"
            "Попробуйте позже. Если повторяется — смотри logs/bot.log.",
            reply_markup=main_menu_kb(),
        )
        await state.clear()


def _count_plan_plates(tracks: list) -> Counter:
    counts: Counter = Counter()
    for track in tracks or []:
        for item in track.get("items") or []:
            if not item:
                continue
            length = round(float(item.get("length") or 0), 2)
            width_raw = item.get("width") or item.get("main_w") or 1.2
            if isinstance(width_raw, (int, float)) and 0 < width_raw < 10:
                width_mm = int(round(float(width_raw) * 1000))
            else:
                width_mm = int(round(float(width_raw or 1200)))
            load_code = cfg.normalize_load_code(item.get("load_code", 8))
            if length > 0 and width_mm > 0:
                counts[(length, width_mm, load_code)] += 1
            for sec in item.get("secondary_cuts") or []:
                if not sec:
                    continue
                sec_length = round(float(sec.get("target_length") or length or 0), 2)
                sec_width_raw = sec.get("width", 0)
                if isinstance(sec_width_raw, (int, float)) and 0 < sec_width_raw < 10:
                    sec_width = int(round(float(sec_width_raw) * 1000))
                else:
                    sec_width = int(round(float(sec_width_raw or 0)))
                sec_load = cfg.normalize_load_code(sec.get("load_code", load_code))
                if sec_length > 0 and sec_width > 0:
                    counts[(sec_length, sec_width, sec_load)] += 1
    return counts


def _log_plan_vs_db_mismatch(
    kp_ids_in_production: list[int],
    plan_plates: Counter,
) -> None:
    from bot.services import kp_persistence as kp_db

    try:
        db_plates: Counter = Counter()
        with kp_db._connect(kp_db.DEFAULT_DB) as conn:
            cur = conn.cursor()
            placeholders = ",".join("?" * len(kp_ids_in_production))
            cur.execute(
                f"""
                SELECT length_m, width_m, load_class, SUM(qty)
                FROM kp_plates
                WHERE kp_id IN ({placeholders})
                AND status IN ('в производстве', 'в плане')
                GROUP BY length_m, width_m, load_class
                """,
                kp_ids_in_production,
            )
            for length_m, width_m, load_class, qty in cur.fetchall():
                load_code = cfg.normalize_load_code(load_class // 100)
                width_mm = int(round(width_m * 1000))
                db_plates[(round(length_m, 2), width_mm, load_code)] += qty
        extra_in_plan = plan_plates - db_plates
        missing_in_plan = db_plates - plan_plates
        if extra_in_plan:
            logger.error(
                "[ПЛАН] В плане есть плиты, которых нет в БД (придуманные): %s",
                dict(extra_in_plan),
            )
        if missing_in_plan:
            logger.warning(
                "[ПЛАН] В БД есть плиты, не попавшие в план (остатки?): %s",
                dict(missing_in_plan),
            )
    except Exception as exc:
        logger.exception("[ПЛАН] Проверка план vs БД: %s", exc)
