"""Обработчики создания коммерческих предложений PDF/XLSX"""
import asyncio
import os
import re
import sys
import math
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Any
from collections import defaultdict

logger = logging.getLogger(__name__)

from aiogram import Router, F
from aiogram.types import Message, FSInputFile, CallbackQuery
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext

# Добавляем корень проекта в sys.path
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.config_and_data import set_plate_lists_from_text, get_current_plate_order, PlateOrder, length_dm_to_m
import core.config_and_data as cfg
from core.commercial_offer import generate_commercial_offer_pdf
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.kp_plate_weight import resolve_kp_line_weight_kg
from core.db_config import PB_DB_PATH
from core.kp_offer_utils import (
    append_transport_to_order_data,
    format_offer_quantity,
    get_offer_item_unit,
    is_transport_offer_item,
)
from core.reinforcement_db import get_reinforcement
from core.visualization import visualize_plan
from core import kp_db
from core.exceptions import PlateParseError, FileGenerationError

# Настройка логирования
logger = logging.getLogger(__name__)

from ..keyboards import main_menu_kb, conditions_choice_kb, save_to_db_kb, save_to_db_with_files_kb, cancel_process_kb, managers_selection_kb, confirm_plates_list_kb, wide_plates_actions_kb, transport_choice_kb
from ..states import KPStates
from ..bot_config import OUTPUTS_DIR_STR

router = Router()

# Лимит длины сообщения Telegram
MAX_MESSAGE_LEN = 4096

# Кэш заказов пользователей
ORDER_CACHE: Dict[int, list] = {}

# Кеш результата оптимизации по user_id (для отложенной генерации схемы)
OPT_PLAN_CACHE: Dict[int, dict] = {}


async def _enter_kp_manager_selection(message: Message, state: FSMContext) -> bool:
    """
    Переход к выбору менеджера (шаг 2 из 5).
    Возвращает True, если список менеджеров отправлен; иначе сообщает об ошибке и очищает state.
    """
    from core.kp_db import get_all_managers

    await state.update_data(kp_accumulated_plates_text="", kp_accumulated_initial_lines=[])

    managers = get_all_managers()
    if not managers:
        await message.answer(
            "⚠️ В базе данных нет менеджеров.\n"
            "Обратитесь к администратору.",
            reply_markup=main_menu_kb(),
        )
        await state.clear()
        return False

    await state.set_state(KPStates.waiting_manager_selection)
    await message.answer(
        "📄 Создание коммерческого предложения\n\n"
        "Шаг 2 из 5: Выберите менеджера",
        reply_markup=managers_selection_kb(managers),
    )
    return True


async def _send_plates_preview_xlsx(
    message: Message,
    *,
    plates_text: str,
    initial_user_plate_lines: list[str],
    forced_wide_line_indexes: list[int] | None = None,
) -> bool:
    """
    Строит XLSX превью списка плит и отправляет документом.
    Возвращает True при успехе; при ошибке логирует и возвращает False.
    """
    try:
        from core.plates_preview_xlsx import build_plates_reconciliation_preview_xlsx

        preview_name = (
            f"Превью_списка_плит_{message.from_user.id}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        preview_path = os.path.join(OUTPUTS_DIR_STR, preview_name)
        await asyncio.to_thread(
            build_plates_reconciliation_preview_xlsx,
            preview_path,
            plates_text=plates_text,
            initial_user_plate_lines=initial_user_plate_lines,
            forced_wide_line_indexes=forced_wide_line_indexes,
        )
        await message.answer_document(
            FSInputFile(preview_path),
            caption="📊 Сверка строк: ввод → распознано → как в КП",
        )
        return True
    except Exception as preview_exc:
        logger.exception(
            "[COMMERCIAL] Не удалось сформировать XLSX превью списка плит: %s",
            preview_exc,
        )
        return False


async def _prompt_kp_step1_plates(message: Message) -> None:
    """Отправляет пользователю приглашение к шагу 1 (ввод списка плит)."""
    await message.answer(
        "📄 Создание коммерческого предложения\n\n"
        "Шаг 1 из 5: Пришлите список плит в свободной форме\n\n"
        "Примеры форматов:\n"
        "• '1.2×3.39 — 2 шт'\n"
        "• '0.32×6.63 — 4 шт'\n"
        "• 'ПБ 38-12-8п 2'\n"
        "• 'ПБ 66-3-8п 4'",
        reply_markup=main_menu_kb(),
    )
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb(),
    )


async def _prompt_transport_choice(message: Message, state: FSMContext, discount_percent: float) -> None:
    """Запрашивает, нужно ли добавить транспортные расходы в КП."""
    await state.set_state(KPStates.waiting_transport_choice)
    await message.answer(
        f"✅ Скидка: {discount_percent}%\n\n"
        "Шаг 5 из 6: Транспортные расходы\n\n"
        "Добавить транспортные расходы в КП?",
        reply_markup=transport_choice_kb(),
    )


async def _prompt_conditions_choice(message: Message, state: FSMContext) -> None:
    """Запрашивает условия поставки и оплаты после ввода транспорта."""
    await state.set_state(KPStates.waiting_conditions_choice)
    await message.answer(
        "Шаг 6 из 6: Условия поставки и оплаты\n\n"
        "Выберите вариант:",
        reply_markup=conditions_choice_kb(),
    )


@router.message(F.text == "📝 Создать КП")
@router.message(Command("commercial_offer"))
async def btn_commercial_offer(message: Message, state: FSMContext):
    """Старт сценария КП: шаг 1 — список плит (менеджеры проверяются заранее)."""
    # Проверяем менеджеров до входа в сценарий (чтобы не пройти плиты и упереться в пустой список)
    from core.kp_db import get_all_managers

    if not get_all_managers():
        await message.answer(
            "⚠️ В базе данных нет менеджеров.\n"
            "Обратитесь к администратору.",
            reply_markup=main_menu_kb(),
        )
        return

    await state.set_state(KPStates.waiting_plates_list)
    await state.update_data(kp_accumulated_plates_text="", kp_accumulated_initial_lines=[])
    await _prompt_kp_step1_plates(message)


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
        "Шаг 3 из 5: Введите имя клиента\n"
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
    """Шаг 3: Получаем имя клиента, затем запрос скидки (шаг 4)."""
    client_name = message.text.strip()

    await state.update_data(client_name=client_name)

    await state.set_state(KPStates.waiting_discount)
    await message.answer(
        f"✅ Клиент: {client_name}\n\n{_build_discount_request_text()}",
        reply_markup=main_menu_kb(),
    )
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb(),
    )


@router.message(KPStates.waiting_plates_list)
async def receive_plates_list(message: Message, state: FSMContext):
    """Шаг 1: Получаем список плит (текст или фото), показываем список и кнопку «Подтвердить»"""

    from core.plate_text_normalizer import get_wide_plate_lines, normalize_order_text

    is_photo = False
    raw_plate_lines: list[str] = []
    ocr_plates_snapshot: list[dict[str, Any]] = []
    initial_user_plate_lines: list[str] = []
    ocr_raw_text: str = ""

    if message.photo:
        is_photo = True
        await message.answer("📸 Фото получено! Распознаю через GPT-4o...")

        photo = message.photo[-1]
        user_id = message.from_user.id
        os.makedirs("tmp", exist_ok=True)
        photo_path = os.path.join("tmp", f"{user_id}_commercial_photo.jpg")

        try:
            await message.bot.download(photo, destination=photo_path)

            from core.ocr_gpt import recognize_text_smart
            recognition_mode = os.getenv("OCR_RECOGNITION_MODE", "full_gpt").strip().lower()
            if recognition_mode not in {"full_gpt", "hybrid"}:
                recognition_mode = "full_gpt"
            result = await recognize_text_smart(
                photo_path,
                force_gpt=(recognition_mode == "full_gpt"),
                show_cost=True,
                mode=recognition_mode,  # full_gpt (по умолчанию) или hybrid
            )

            if result and result.get('text'):
                plates_text = result['text']
                ocr_raw_text = (result.get("text") or "").strip()
                initial_user_plate_lines = [
                    ln.strip()
                    for ln in re.split(r"[\n;]+", ocr_raw_text)
                    if ln.strip()
                ]
                ocr_plates_snapshot = list(result.get("plates") or [])
                cost = result.get('cost_usd', 0)
                recognized_count = sum(p.get('qty', 1) for p in result.get('plates', []))

                cost_note = ""
                if cost > 0:
                    rub_cost = cost * 75
                    cost_note = f"\n\n💰 Стоимость распознавания: ${cost:.4f} (~{rub_cost:.2f}₽)"

                await message.answer(f"✅ Распознано через GPT-4o Vision{cost_note}")
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
            logger.exception(f"[COMMERCIAL] Ошибка распознавания фото: {e}")
            await message.answer(
                f"❌ Ошибка при обработке фото: {str(e)}\n\n"
                "Попробуйте прислать список текстом."
            )
            return

    elif message.text:
        plates_text = message.text.strip()
        raw_plate_lines = [l.strip() for l in re.split(r"[\n;]+", plates_text) if l.strip()]
        initial_user_plate_lines = list(raw_plate_lines)
        recognized_count = 0

        # Парсим текст для подсчёта количества и проверки широких плит
        try:
            set_plate_lists_from_text(plates_text)
            recognized_count = sum(get_current_plate_order().plate_load_details.values())
        except PlateParseError as e:
            logger.warning(f"[COMMERCIAL] Ошибка парсинга списка плит от пользователя {message.from_user.id}: {e}")
            await message.answer(
                f"❌ Не удалось распознать список плит:\n{e}\n\n"
                "💡 Проверьте формат:\n"
                "• ПБ 78-12-8п 5 шт\n"
                "• 1.2×3.39 — 2\n"
                "• 0,32x6,63 - 4",
                reply_markup=main_menu_kb()
            )
            await message.answer(
                "Или нажмите кнопку ниже для отмены:",
                reply_markup=cancel_process_kb()
            )
            return

    else:
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

    # Нормализуем к каноническому виду (как парсим) для отображения и хранения
    norm = normalize_order_text(plates_text)
    plates_text_to_store = norm.normalized_text.strip() if norm.normalized_text.strip() else plates_text

    data_before_merge = await state.get_data()
    acc_text = (data_before_merge.get("kp_accumulated_plates_text") or "").strip()
    acc_lines: list[str] = list(data_before_merge.get("kp_accumulated_initial_lines") or [])
    if acc_text:
        merged_raw = acc_text + "\n" + plates_text_to_store
        norm_merged = normalize_order_text(merged_raw)
        plates_text_to_store = (
            norm_merged.normalized_text.strip()
            if norm_merged.normalized_text.strip()
            else merged_raw.strip()
        )
        initial_user_plate_lines = acc_lines + initial_user_plate_lines

    try:
        set_plate_lists_from_text(plates_text_to_store)
        recognized_count = sum(get_current_plate_order().plate_load_details.values())
    except PlateParseError as merge_parse_exc:
        logger.warning(
            "[COMMERCIAL] Парсинг объединённого списка плит: %s",
            merge_parse_exc,
        )
        recognized_count = 0

    # Проверяем наличие плит с шириной > 12 дм (по нормализованному тексту)
    wide_lines = get_wide_plate_lines(plates_text_to_store)
    wide_plate_lines: list[str] = [line for line, _qty in wide_lines]

    # Формируем текст для отображения — показываем список в том виде, как мы его парсим
    display_limit = MAX_MESSAGE_LEN - 200
    if len(plates_text_to_store) > display_limit:
        display_text = plates_text_to_store[:display_limit] + f"\n… (полный список сохранён, всего {len(plates_text_to_store)} символов)"
    else:
        display_text = plates_text_to_store

    # Формируем сообщение со списком
    if is_photo:
        count_note = f"\n\nРаспознано позиций: {recognized_count} шт." if recognized_count > 0 else ""
        list_msg = f"📋 Распознанный список плит:\n\n{display_text}{count_note}"
    else:
        list_msg = (
            f"📋 Список плит:\n\n{display_text}\n\n"
            f"Количество в заказе: {recognized_count}. "
            f"Распознано: {recognized_count}. Одинаковое: да."
        )

    # Сохраняем данные в state (нормализованный текст — как мы парсим)
    await state.update_data(
        plates_text=plates_text_to_store,
        recognized_count=recognized_count,
        wide_plate_lines=wide_plate_lines,
        forced_wide_line_indexes=[],
        replacement_done=False,
        plates_preview_sent=False,
        is_photo=is_photo,
        raw_plate_lines=raw_plate_lines,
        ocr_plates_snapshot=ocr_plates_snapshot,
        initial_user_plate_lines=initial_user_plate_lines,
        ocr_raw_text=ocr_raw_text,
    )
    await state.set_state(KPStates.waiting_plates_confirm)

    await message.answer(list_msg, reply_markup=confirm_plates_list_kb())
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
    )


def _build_discount_request_text() -> str:
    return (
        "Шаг 4 из 5: Введите процент скидки\n"
        "(Просто число, например: 0, 5, 10, 15)\n"
        "0 = без скидки"
    )


def _build_wide_plates_replacement_example(wide_plate_lines: list[str]) -> str:
    """
    Формирует пример замены для плит шириной >12 дм:
    ПБ L-15-8п 2 -> ПБ L-12-8п 2 + ПБ L-3,0-8п 2.
    """
    pb_line_re = re.compile(
        r"(?i)\b(?P<prefix>п[бк])\s*"
        r"(?P<length>[\d,.]+)\s*-\s*"
        r"(?P<width>[\d,.]+)\s*-\s*"
        r"(?P<load>[\d,.]+)\s*п?\s*"
        r"(?P<qty>\d+)?\s*(?:шт\.?|штук)?\s*$"
    )

    example_blocks: list[str] = []
    for raw_line in wide_plate_lines:
        line = raw_line.strip()
        match = pb_line_re.search(line)
        if not match:
            continue

        prefix = match.group("prefix").upper()
        length_part = match.group("length").replace(".", ",")
        load_part = match.group("load").replace(".", ",")
        qty = (match.group("qty") or "").strip()
        qty_suffix = f" {qty}" if qty else ""

        example_blocks.append(
            f"• {prefix} {length_part}-12-{load_part}п{qty_suffix}\n"
            f"• {prefix} {length_part}-3,0-{load_part}п{qty_suffix}"
        )

    if not example_blocks:
        return "Пришлите список плит, на которые их нужно заменить."

    examples_block = "\n\n".join(example_blocks)
    return (
        "ПРИМЕР\n"
        "Пришлите список плит, на которые их нужно заменить.\n"
        "_____________________________________________________\n"
        "Пример из ваших плит:\n\n"
        f"{examples_block}\n"
        "КОНЕЦ ПРИМЕРА"
    )


@router.callback_query(F.data == "confirm_plates_list", KPStates.waiting_plates_confirm)
async def confirm_plates_list_callback(callback: CallbackQuery, state: FSMContext):
    """Обработчик нажатия «Подтвердить» для списка плит"""
    data = await state.get_data()
    wide_plate_lines: list[str] = data.get("wide_plate_lines", [])
    replacement_done: bool = data.get("replacement_done", False)
    plates_preview_sent: bool = data.get("plates_preview_sent", False)

    await callback.answer()

    if replacement_done:
        await state.update_data(
            wide_plate_lines=[],
            forced_wide_line_indexes=[],
            replacement_done=False,
            plates_preview_sent=False,
        )
        await _enter_kp_manager_selection(callback.message, state)
        return

    if wide_plate_lines:
        await state.update_data(plates_preview_sent=False)
        wide_list_text = "\n".join(f"• {line}" for line in wide_plate_lines)
        example_text = _build_wide_plates_replacement_example(wide_plate_lines)
        await state.set_state(KPStates.waiting_wide_plates_replacement)
        await callback.message.answer(
            f"⚠️ В списке есть плиты шириной больше 12 дм:\n{wide_list_text}\n\n"
            f"{example_text}"
        )
        await callback.message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=wide_plates_actions_kb()
        )
        return

    if not plates_preview_sent:
        plates_text = data.get("plates_text", "")
        initial_lines = list(data.get("initial_user_plate_lines") or [])
        preview_ok = await _send_plates_preview_xlsx(
            callback.message,
            plates_text=plates_text,
            initial_user_plate_lines=initial_lines,
            forced_wide_line_indexes=list(data.get("forced_wide_line_indexes") or []),
        )
        await state.update_data(plates_preview_sent=True)
        if preview_ok:
            await callback.message.answer(
                "📊 Проверьте файл сверки. Если всё верно, нажмите «✅ Подтвердить» "
                "ещё раз для выбора менеджера.\n"
                "Если нужно изменить список, нажмите «🔄 Заменить».\n"
                "Чтобы сохранить список и прислать ещё плит отдельным сообщением — «➕ Продолжить КП».",
                reply_markup=confirm_plates_list_kb(),
            )
        else:
            await callback.message.answer(
                "⚠️ Не удалось сформировать XLSX превью. Нажмите «✅ Подтвердить» ещё раз "
                "для выбора менеджера.\n"
                "Если нужно изменить список, нажмите «🔄 Заменить».\n"
                "Чтобы добавить позиции отдельным вводом — «➕ Продолжить КП».",
                reply_markup=confirm_plates_list_kb(),
            )
        return

    await state.update_data(plates_preview_sent=False)
    await _enter_kp_manager_selection(callback.message, state)


@router.callback_query(F.data == "continue_kp_plates", KPStates.waiting_plates_confirm)
async def continue_kp_plates_callback(callback: CallbackQuery, state: FSMContext):
    """Сохраняет текущий список для КП и возвращает к шагу 1 для ввода следующей порции плит."""
    data = await state.get_data()
    wide_plate_lines: list[str] = data.get("wide_plate_lines", [])

    await callback.answer()

    if wide_plate_lines:
        await state.update_data(plates_preview_sent=False)
        wide_list_text = "\n".join(f"• {line}" for line in wide_plate_lines)
        example_text = _build_wide_plates_replacement_example(wide_plate_lines)
        await state.set_state(KPStates.waiting_wide_plates_replacement)
        await callback.message.answer(
            f"⚠️ В списке есть плиты шириной больше 12 дм:\n{wide_list_text}\n\n"
            f"{example_text}"
        )
        await callback.message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=wide_plates_actions_kb(),
        )
        return

    plates_text = (data.get("plates_text") or "").strip()
    if not plates_text:
        await callback.message.answer(
            "❌ Нет списка плит для сохранения. Пришлите список текстом или фото.",
            reply_markup=cancel_process_kb(),
        )
        return

    initial_lines = list(data.get("initial_user_plate_lines") or [])
    await state.update_data(
        kp_accumulated_plates_text=plates_text,
        kp_accumulated_initial_lines=initial_lines,
        plates_preview_sent=False,
        replacement_done=False,
        wide_plate_lines=[],
        raw_plate_lines=[],
        ocr_plates_snapshot=[],
        ocr_raw_text="",
        plates_text="",
        initial_user_plate_lines=[],
        recognized_count=0,
        is_photo=False,
    )
    await state.set_state(KPStates.waiting_plates_list)

    await callback.message.answer(
        "📌 Текущий список плит сохранён для КП. Пришлите ещё позиции — они будут объединены "
        "с уже сохранёнными в следующем превью.\n\n"
        "Шаг 1 из 5: дополнительный список плит (текст или фото):",
        reply_markup=main_menu_kb(),
    )
    await callback.message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb(),
    )


@router.callback_query(F.data == "replace_plates_list", KPStates.waiting_plates_confirm)
async def replace_plates_list_callback(callback: CallbackQuery, state: FSMContext):
    """Возврат к шагу ввода списка плит с полным сбросом текущих данных."""
    await callback.answer()
    await state.clear()
    await state.set_state(KPStates.waiting_plates_list)
    await _prompt_kp_step1_plates(callback.message)


@router.message(KPStates.waiting_wide_plates_replacement)
async def receive_wide_plates_replacement(message: Message, state: FSMContext):
    """Получаем список замен для плит шириной > 12 дм и формируем итоговый список"""
    import re as _re

    from core.plate_text_normalizer import get_wide_plate_lines, normalize_order_text

    if not message.text:
        await message.answer(
            "❌ Пришлите список замен текстом.",
            reply_markup=cancel_process_kb()
        )
        return

    replacement_text = message.text.strip()

    # Валидируем список замен через парсер
    try:
        set_plate_lists_from_text(replacement_text)
    except PlateParseError as e:
        logger.warning(f"[COMMERCIAL] Ошибка парсинга списка замен: {e}")
        await message.answer(
            f"❌ Не удалось распознать список замен:\n{e}\n\n"
            "💡 Проверьте формат:\n"
            "• ПБ 78-12-8п 5 шт\n"
            "• 1.2×3.39 — 2\n\n"
            "Пришлите исправленный список:"
        )
        return

    data = await state.get_data()
    original_plates_text: str = data.get("plates_text", "")
    wide_plate_lines: list[str] = data.get("wide_plate_lines", [])

    # Нормализуем список замен к каноническому виду
    norm_rep = normalize_order_text(replacement_text)
    replacement_lines = norm_rep.normalized_lines if norm_rep.normalized_lines else [l.strip() for l in _re.split(r"[\n;]+", replacement_text) if l.strip()]

    # Слияние: заменяем строки с шириной > 12 дм на строки из списка замен.
    # Для превью колонки A сохраняем исходную wide-строку только в первой строке
    # её блока, последующие строки блока оставляем пустыми (через строку).
    original_lines = [l.strip() for l in _re.split(r"[\n;]+", original_plates_text) if l.strip()]

    wide_set = set(wide_plate_lines)
    original_wide_lines = [line for line in original_lines if line in wide_set]
    wide_count = len(original_wide_lines)
    replacements_by_wide: list[list[str]] = []
    if wide_count > 0:
        if len(replacement_lines) >= wide_count:
            base = len(replacement_lines) // wide_count
            rem = len(replacement_lines) % wide_count
            cursor = 0
            for wide_idx in range(wide_count):
                take = base + (1 if wide_idx < rem else 0)
                replacements_by_wide.append(replacement_lines[cursor:cursor + take])
                cursor += take
        else:
            # Замен меньше, чем wide-строк: даём по одной замене по порядку, остальным пусто.
            for wide_idx in range(wide_count):
                if wide_idx < len(replacement_lines):
                    replacements_by_wide.append([replacement_lines[wide_idx]])
                else:
                    replacements_by_wide.append([])

    replacement_wide_idx = 0
    merged_lines: list[str] = []
    merged_is_wide: list[bool] = []
    preview_user_plate_lines: list[str] = []

    for line in original_lines:
        if line in wide_set and replacement_wide_idx < len(replacements_by_wide):
            block_replacements = replacements_by_wide[replacement_wide_idx]
            replacement_wide_idx += 1
            if block_replacements:
                for idx_in_block, rep_line in enumerate(block_replacements):
                    merged_lines.append(rep_line)
                    merged_is_wide.append(True)
                    preview_user_plate_lines.append(line if idx_in_block == 0 else "")
            else:
                # Если для wide-строки не прислали замену — сохраняем исходную строку.
                merged_lines.append(line)
                merged_is_wide.append(True)
                preview_user_plate_lines.append(line)
        else:
            merged_lines.append(line)
            merged_is_wide.append(False)
            preview_user_plate_lines.append(line)

    consumed_replacements = sum(len(block) for block in replacements_by_wide)
    replacement_index = consumed_replacements
    # Если замен прислали больше, чем wide-строк/распределение — добавляем остаток в конец.
    # Эти строки считаем частью wide-блока; в A показываем только первую из них.
    while replacement_index < len(replacement_lines):
        merged_lines.append(replacement_lines[replacement_index])
        merged_is_wide.append(True)
        preview_user_plate_lines.append("")
        replacement_index += 1

    final_plates_text_raw = "\n".join(merged_lines)
    # Итоговый список в каноническом виде
    norm_final = normalize_order_text(final_plates_text_raw)
    final_plates_text = norm_final.normalized_text.strip() if norm_final.normalized_text.strip() else final_plates_text_raw

    # Считаем количество по итоговому списку
    try:
        set_plate_lists_from_text(final_plates_text)
        final_count = sum(get_current_plate_order().plate_load_details.values())
    except PlateParseError:
        final_count = 0

    # Формируем текст для отображения (учитываем лимит)
    display_limit = MAX_MESSAGE_LEN - 200
    if len(final_plates_text) > display_limit:
        display_text = final_plates_text[:display_limit] + f"\n… (полный список сохранён, всего {len(final_plates_text)} символов)"
    else:
        display_text = final_plates_text

    count_note = (
        f"\n\nКоличество в заказе: {final_count}. "
        f"Распознано: {final_count}. Одинаковое: да."
    ) if final_count > 0 else ""

    list_msg = f"📋 Итоговый список плит (с заменами):\n\n{display_text}{count_note}"

    if not preview_user_plate_lines:
        preview_user_plate_lines = list(original_lines)
    forced_wide_line_indexes = [idx for idx, is_wide in enumerate(merged_is_wide) if is_wide]
    await state.update_data(
        plates_text=final_plates_text,
        wide_plate_lines=[],
        replacement_done=True,
        plates_preview_sent=False,
        raw_plate_lines=merged_lines,
        initial_user_plate_lines=preview_user_plate_lines,
        forced_wide_line_indexes=forced_wide_line_indexes,
    )
    await state.set_state(KPStates.waiting_plates_confirm)

    await message.answer(list_msg, reply_markup=confirm_plates_list_kb())

    await _send_plates_preview_xlsx(
        message,
        plates_text=final_plates_text,
        initial_user_plate_lines=preview_user_plate_lines,
        forced_wide_line_indexes=forced_wide_line_indexes,
    )

    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
    )


@router.callback_query(F.data == "skip_wide_plates", KPStates.waiting_wide_plates_replacement)
async def skip_wide_plates_callback(callback: CallbackQuery, state: FSMContext):
    """Пропускает все строки с плитами шире 12 дм (исключает их из заказа)."""
    await callback.answer()

    import re as _re
    from core.plate_text_normalizer import normalize_order_text

    data = await state.get_data()
    original_plates_text: str = data.get("plates_text", "")
    wide_plate_lines: list[str] = list(data.get("wide_plate_lines", []) or [])

    original_lines = [l.strip() for l in _re.split(r"[\n;]+", original_plates_text) if l.strip()]
    wide_set = set(wide_plate_lines)
    filtered_lines = [line for line in original_lines if line not in wide_set]
    skipped_count = len(original_lines) - len(filtered_lines)

    if not filtered_lines:
        await callback.message.answer(
            "⚠️ После пропуска широких плит список стал пустым.\n"
            "Пришлите список замен или нажмите «🔄 Заменить» и введите новый список.",
            reply_markup=cancel_process_kb(),
        )
        return

    final_plates_text_raw = "\n".join(filtered_lines)
    norm_final = normalize_order_text(final_plates_text_raw)
    final_plates_text = (
        norm_final.normalized_text.strip() if norm_final.normalized_text.strip() else final_plates_text_raw
    )

    try:
        set_plate_lists_from_text(final_plates_text)
        final_count = sum(get_current_plate_order().plate_load_details.values())
    except PlateParseError:
        final_count = 0

    display_limit = MAX_MESSAGE_LEN - 200
    if len(final_plates_text) > display_limit:
        display_text = final_plates_text[:display_limit] + f"\n… (полный список сохранён, всего {len(final_plates_text)} символов)"
    else:
        display_text = final_plates_text

    count_note = (
        f"\n\nКоличество в заказе: {final_count}. "
        f"Распознано: {final_count}. Одинаковое: да."
    ) if final_count > 0 else ""

    list_msg = (
        f"📋 Итоговый список плит (широкие пропущены: {skipped_count}):\n\n"
        f"{display_text}{count_note}"
    )

    updated_user_plate_lines = list(filtered_lines)
    await state.update_data(
        plates_text=final_plates_text,
        wide_plate_lines=[],
        replacement_done=True,
        plates_preview_sent=False,
        raw_plate_lines=filtered_lines,
        initial_user_plate_lines=updated_user_plate_lines,
        forced_wide_line_indexes=[],
    )
    await state.set_state(KPStates.waiting_plates_confirm)

    await callback.message.answer(list_msg, reply_markup=confirm_plates_list_kb())

    await _send_plates_preview_xlsx(
        callback.message,
        plates_text=final_plates_text,
        initial_user_plate_lines=updated_user_plate_lines,
        forced_wide_line_indexes=[],
    )

    await callback.message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
    )


@router.message(KPStates.waiting_plates_confirm)
async def plates_confirm_unexpected_text(message: Message, state: FSMContext):
    """Подсказка пользователю нажать кнопку под списком плит."""
    await message.answer(
        "Нажмите «✅ Подтвердить», «🔄 Заменить» или «➕ Продолжить КП» под списком плит, "
        "или «◀️ Назад в меню» для отмены."
    )


@router.message(KPStates.waiting_discount)
async def receive_discount_and_ask_conditions(message: Message, state: FSMContext):
    """Шаг 4 из 6: получаем процент скидки и переходим к выбору транспорта."""
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
    await state.update_data(
        discount_percent=discount_percent,
        transport_hours=None,
        transport_price_per_hour=None,
    )
    await _prompt_transport_choice(message, state, discount_percent)


@router.callback_query(KPStates.waiting_transport_choice, F.data.in_(["transport_add", "transport_skip"]))
async def receive_transport_choice(callback: CallbackQuery, state: FSMContext):
    """Шаг 5 из 6: выбор, добавлять ли транспорт в КП."""
    choice = callback.data
    await callback.answer()

    if choice == "transport_add":
        await callback.message.edit_text("✅ Выбрано: Добавить транспорт")
        await state.set_state(KPStates.waiting_transport_hours)
        await callback.message.answer(
            "Введите количество часов транспортных услуг.\n"
            "Пример: 2 или 2,5",
            reply_markup=main_menu_kb()
        )
        await callback.message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return

    if choice == "transport_skip":
        await state.update_data(transport_hours=None, transport_price_per_hour=None)
        await callback.message.edit_text("✅ Выбрано: Без транспорта")
        await _prompt_conditions_choice(callback.message, state)


@router.message(KPStates.waiting_transport_hours)
async def receive_transport_hours(message: Message, state: FSMContext):
    """Шаг 5 из 6: получаем количество часов транспорта."""
    try:
        transport_hours = float(message.text.strip().replace(',', '.'))
        if transport_hours <= 0:
            raise ValueError
    except ValueError:
        await message.answer(
            "❌ Введите число больше 0. Например: 2 или 2,5.",
            reply_markup=main_menu_kb()
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return

    await state.update_data(transport_hours=transport_hours)
    await state.set_state(KPStates.waiting_transport_price)
    await message.answer(
        f"✅ Часы транспорта: {format_offer_quantity(transport_hours)}\n\n"
        "Теперь введите цену транспортных услуг за 1 час.",
        reply_markup=main_menu_kb()
    )
    await message.answer(
        "Или нажмите кнопку ниже для отмены:",
        reply_markup=cancel_process_kb()
    )


@router.message(KPStates.waiting_transport_price)
async def receive_transport_price_and_ask_conditions(message: Message, state: FSMContext):
    """Шаг 5 из 6: получаем цену транспорта и переходим к условиям."""
    try:
        transport_price_per_hour = float(message.text.strip().replace(',', '.'))
        if transport_price_per_hour <= 0:
            raise ValueError
    except ValueError:
        await message.answer(
            "❌ Введите цену за час числом больше 0. Например: 9800 или 9800,50.",
            reply_markup=main_menu_kb()
        )
        await message.answer(
            "Или нажмите кнопку ниже для отмены:",
            reply_markup=cancel_process_kb()
        )
        return

    data = await state.get_data()
    transport_hours = data.get("transport_hours")
    await state.update_data(transport_price_per_hour=transport_price_per_hour)
    await message.answer(
        f"✅ Транспорт: {format_offer_quantity(transport_hours)} час. × "
        f"{transport_price_per_hour:,.2f} ₽",
        reply_markup=main_menu_kb()
    )
    await _prompt_conditions_choice(message, state)


@router.callback_query(KPStates.waiting_conditions_choice, F.data.in_(["conditions_default", "conditions_custom"]))
async def receive_conditions_choice(callback: CallbackQuery, state: FSMContext):
    """Шаг 6 из 6: выбор условий (по умолчанию или свои) — inline-кнопки."""
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
    raw_plate_lines: list[str] = list(data.get("raw_plate_lines") or [])
    is_photo_kp: bool = bool(data.get("is_photo", False))
    discount_percent = data.get('discount_percent', 0)
    transport_hours = data.get('transport_hours')
    transport_price_per_hour = data.get('transport_price_per_hour')
    delivery_conditions = data.get('delivery_conditions', '')
    payment_conditions = data.get('payment_conditions', '')
    
    # Показываем сводку
    summary_text = (
        f"📋 Сводка:\n"
        f"• Менеджер: {manager_name}\n"
        f"• Клиент: {client_name}\n"
        f"• Скидка: {discount_percent}%\n"
    )
    if transport_hours and transport_price_per_hour:
        transport_total = float(transport_hours) * float(transport_price_per_hour)
        summary_text += (
            f"• Транспорт: {format_offer_quantity(transport_hours)} час. × "
            f"{float(transport_price_per_hour):,.2f} ₽ = {transport_total:,.2f} ₽\n"
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
        try:
            unparsed_lines, line_contributions, _line_plate_load_details = set_plate_lists_from_text(
                plates_text
            )
        except PlateParseError as e:
            # Ошибка парсинга - показываем понятное сообщение
            logger.warning(f"Ошибка парсинга заказа от пользователя {message.from_user.id}: {e}")
            await message.answer(
                f"❌ Не удалось распознать заказ:\n{e}\n\n"
                f"💡 Проверьте формат:\n"
                f"• ПБ 78-12-8п 5 шт\n"
                f"• 1.2x3.39 - 2\n"
                f"• 0,32x6,63 - 4",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        order = get_current_plate_order()
        await state.update_data(plate_order=order.to_dict())
        
        if unparsed_lines:
            # Показываем только первые 5 нераспознанных строк, чтобы не спамить
            warn_text = "⚠️ Некоторые строки не распознаны:\n"
            warn_text += "\n".join(f"• {line}" for line in unparsed_lines[:5])
            if len(unparsed_lines) > 5:
                warn_text += f"\n... и ещё {len(unparsed_lines) - 5} строк"
            warn_text += (
                "\n\n💡 Я понимаю такие форматы:\n"
                "• ПБ 78-12-8п 3 шт\n"
                "• 1.2×3.39 — 2\n"
                "• 0,32x6,63 - 4\n"
                "• ПБ 66,2-12-8п 6"
            )
            await message.answer(warn_text)
        
        # ✅ ЗАПУСКАЕМ ОПТИМИЗАЦИЮ (упрощённая версия, без разделения по армированию)
        await message.answer("🔄 Оптимизирую раскрой плит для минимизации стоимости...")
        
        # Собираем все плиты в один список orders_2d из заказа (изоляция по пользователю)
        orders_2d = order.to_orders_2d()
        if orders_2d:
            logger.info("[COMMERCIAL] Используем PlateOrder (с нагрузками)")
        
        if orders_2d:
            logger.info(
                f"[COMMERCIAL] Всего плит для оптимизации: {sum(o['qty'] for o in orders_2d)} шт, типов: {len(orders_2d)}"
            )
            
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
                    
                    # Кешируем для отложенной генерации схемы (по нажатию кнопки)
                    OPT_PLAN_CACHE[message.from_user.id] = {
                        'plan': optimization_result,
                        'by_load': optimization.OPT_CASCADING_PLAN_BY_LOAD,
                        'load_map': dict(optimization.LOAD_TO_REINFORCEMENT_MAP),
                    }
                    
                    total_plates = optimization_result.get('total_plates', 0)
                    total_cost = optimization_result.get('total_cost', 0)
                    logger.info(
                        f"[COMMERCIAL] Оптимизация завершена: {total_plates} плит, {total_cost:,} ₽".replace(",", " ")
                    )
                    await message.answer(f"✅ Оптимизация завершена! Использовано {total_plates} исходных плит")
                    
            except Exception as e:
                logger.exception(f"[COMMERCIAL] Ошибка оптимизации: {e}")
                # Продолжаем без оптимизации (цены будут посчитаны по старой логике)
        
        # Используем build_price_rows для получения правильных цен
        from viz_modules.procurement import build_price_rows, build_component_breakdown, build_procurement_items
        from viz_modules.price_utils import load_price_table_from_xlsx
        
        # Загружаем таблицу цен
        price_table = load_price_table_from_xlsx(str(cfg.PRICE_XLSX_PATH))
        
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
        
        # Формируем order_data из того же источника, что и смета (build_procurement_items), чтобы не терять плиты
        order_data = []
        procurement_items = build_procurement_items()
        for item in procurement_items:
            length_m = item['length']
            width_m = item['width']
            qty = item['qty']
            load_code = item.get('load_code')
            if load_code is None:
                load_code = cfg.get_load_code_for_plate(length_m, width_m, default=(6 if width_m < 1.0 else 8))
            matching_row = None
            for row in price_rows:
                if len(row) < 8:
                    continue
                row_name = row[1]
                row_qty = row[2]
                parsed_length, parsed_width = cfg.parse_name_to_sizes(row_name)
                if parsed_length is None or parsed_width is None:
                    continue
                parsed_load = cfg.parse_load_code_from_name(row_name)
                # Сопоставление по длине с допуском 0.01 м: различаем 6.35 и 6.37 (не склеиваем в один round(63.5)=64).
                # Раньше int(round(parsed_length*10))==int(round(length_m*10)) давало одну строку сметы на 6.35 и 6.37.
                # Для 12,5п: из имени парсится 13, в заказе 12.5 — сравниваем через load_code_for_price_match
                length_match = abs(parsed_length - length_m) < 0.01
                if (length_match
                        and abs(parsed_width - width_m) < 0.01
                        and cfg.load_code_for_price_match(parsed_load) == cfg.load_code_for_price_match(load_code)):
                    matching_row = row
                    break
            if matching_row:
                name = matching_row[1]
                price_str = matching_row[7]
                try:
                    unit_price = float(price_str.replace(' ', '').replace(',', '.'))
                except (ValueError, AttributeError):
                    unit_price = 0.0
                match = re.search(r'ПБ\s+([\d,]+)-', name)
                length_dm_raw = match.group(1).strip() if match else ''
                # Не подменять имя/длину, если у позиции явно задан length_dm_raw и он отличается от подстроки в строке сметы (57 vs 57,1)
                item_ldr = item.get('length_dm_raw')
                if item_ldr and length_dm_raw and length_dm_raw != item_ldr:
                    name = cfg.make_plate_name(length_m, width_m, load_code=load_code, length_dm_raw=item_ldr)
                    length_dm_raw = item_ldr
            else:
                # #region agent log
                try:
                    _log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'debug-b59370.log')
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "b59370", "hypothesisId": "H3", "location": "commercial:order_data_no_match", "message": "no matching_row, using make_plate_name", "data": {"length_m": length_m, "width_m": width_m}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
                # #endregion
                # length_dm_raw из item прокинут из build_procurement_items (57,1 не теряется)
                item_ldr = item.get('length_dm_raw') or None
                name = cfg.make_plate_name(length_m, width_m, load_code=load_code, length_dm_raw=item_ldr)
                unit_price = 0.0
                length_dm_raw = item_ldr or (
                    f'{length_m * 10:.1f}'.rstrip('0').rstrip('.').replace('.', ',') if length_m else ''
                )
            # Если кэш номенклатуры уже содержит canonical_name для этой позиции —
            # используем его вместо имени из сметы/make_plate_name, а nomenclature_id
            # подставляем сразу, чтобы enrich не делал повторный поиск по БД.
            cache_name = item.get('canonical_name')
            cache_nid = item.get('nomenclature_id')
            if cache_name:
                name = cache_name
                _raw_match = re.search(r'ПБ\s+([\d,]+)-', name)
                length_dm_raw = _raw_match.group(1).strip() if _raw_match else length_dm_raw
            # Вес КП: plate_weights по размерам (сумма на строку — для save_kp_to_db)
            _, total_weight_kg = resolve_kp_line_weight_kg({
                "length_m": length_m,
                "width_m": width_m,
                "qty": qty,
            })
            # #region agent log (57/57,1: итоговое имя в order_data)
            if 5.69 <= length_m <= 5.73:
                try:
                    _log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'debug-8e9428.log')
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_order_data", "location": "commercial:order_data_loop", "message": "57/57,1: final name in order_data", "data": {"length_m": length_m, "width_m": width_m, "qty": qty, "name": name, "from_matching_row": matching_row[1] if matching_row else None}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            # #region agent log (a9176e: 57/57,1 — order_data: matching_row + итоговое имя)
            if 5.69 <= length_m <= 5.73:
                try:
                    _parsed = (cfg.parse_name_to_sizes(matching_row[1]) if matching_row and len(matching_row) > 1 else (None, None))
                    _log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'debug-a9176e.log')
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "a9176e", "hypothesisId": "H3", "location": "commercial:order_data_loop", "message": "57/57,1 order_data name source", "data": {"length_m": length_m, "qty": qty, "has_matching_row": matching_row is not None, "matching_row_name": matching_row[1] if matching_row and len(matching_row) > 1 else None, "parsed_length": _parsed[0], "item_length_dm_raw": item.get('length_dm_raw'), "cache_name": item.get('canonical_name'), "final_name": name, "final_length_dm_raw": length_dm_raw}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            entry = {
                "name": name,
                "length_m": length_m,
                "length_dm_raw": length_dm_raw,
                "width_m": width_m,
                "qty": qty,
                "load_class": (cfg.normalize_load_code(load_code) or 8) * 100 if load_code is not None else 800,
                "unit_price": unit_price,
                "weight": total_weight_kg,
            }
            if cache_nid is not None:
                entry['nomenclature_id'] = cache_nid
            order_data.append(entry)
        total_order_data = sum(i['qty'] for i in order_data)
        total_orders_2d = sum(o['qty'] for o in orders_2d) if orders_2d else 0
        if total_orders_2d and total_order_data != total_orders_2d:
            logger.error(f"[КП] ПОТЕРЯ ПЛИТ: order_data={total_order_data}, orders_2d={total_orders_2d}")
            await message.answer(
                "⚠️ ВНИМАНИЕ: Обнаружено расхождение в количестве плит!\n"
                f"Заказано: {total_orders_2d} шт\n"
                f"В смете: {total_order_data} шт\n\n"
                "Пожалуйста, проверьте данные и попробуйте заново.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        if not order_data:
            await message.answer(
                "❌ Не удалось распознать ни одной плиты в вашем сообщении.\n"
                "Проверьте формат строк.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
        
        # 🆕 Обогащаем order_data точными названиями из prays_plity
        from core.kp_db import enrich_order_data_with_nomenclature
        order_data = await asyncio.to_thread(enrich_order_data_with_nomenclature, order_data)
        order_data = append_transport_to_order_data(
            order_data,
            transport_hours,
            transport_price_per_hour,
        )

        # #region agent log
        try:
            _od_total = sum(i.get('qty', 0) for i in order_data)
            _od_sample = [(i.get('name', '')[:30], round(i.get('length_m', 0), 3), i.get('qty')) for i in order_data[:3]]
            open(r"c:\Users\Роман\Desktop\Шишов\.cursor\debug.log", "a", encoding="utf-8").write(
                __import__("json").dumps({"hypothesisId": "H3", "location": "commercial:order_data_from_price_rows", "message": "order_data after build_price_rows", "data": {"len_order_data": len(order_data), "total_qty": _od_total, "sample": _od_sample}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n"
            )
        except Exception:
            pass
        # #endregion
        # Сохраняем заказ в кэш
        ORDER_CACHE[message.from_user.id] = order_data
        
        # Генерируем номер и дату КП
        offer_number = f"{message.from_user.id}_{datetime.now().strftime('%Y%m%d%H%M')}"
        offer_date = datetime.now().strftime("%d.%m.%Y")
        
        # 🆕 ПОЛУЧАЕМ СЛЕДУЮЩИЙ НОМЕР КП ИЗ БД
        from core.kp_db import get_next_kp_number
        kp_db_id = get_next_kp_number()
        logger.debug(f"Предполагаемый номер КП из БД: {kp_db_id}")

        try:
            from core.reconciliation_xlsx import build_reconciliation_xlsx

            rec_path = os.path.join(
                OUTPUTS_DIR_STR,
                f"Сверка_КП_{offer_number}_{offer_date.replace('.', '')}.xlsx",
            )
            await asyncio.to_thread(
                build_reconciliation_xlsx,
                rec_path,
                plates_text=plates_text,
                raw_plate_lines=raw_plate_lines,
                unparsed_lines=unparsed_lines,
                line_contributions=line_contributions,
                order_data=order_data,
                is_photo=is_photo_kp,
            )
            await message.answer_document(
                FSInputFile(rec_path),
                caption="📊 Сверка ввода: как прислали → распознано → как в КП",
            )
        except Exception as rec_exc:
            logger.exception("[COMMERCIAL] Не удалось сформировать XLSX сверки: %s", rec_exc)
        
        # Документы генерируются по нажатию кнопок — показываем кнопки сразу
        # Формируем сводку
        plate_items = [item for item in order_data if not is_transport_offer_item(item)]
        total_qty = sum(item['qty'] for item in plate_items)
        summary = f"✅ Коммерческое предложение готово!\n\n"
        summary += f"📋 Заказ:\n"
        for item in order_data:
            summary += (
                f"  • {item['name']} — {format_offer_quantity(item['qty'])} "
                f"{get_offer_item_unit(item)}\n"
            )
        summary += f"\n📊 Всего позиций: {len(order_data)}\n"
        summary += f"📦 Всего плит: {total_qty} шт\n\n"
        summary += "✨ Документы содержат:\n"
        summary += "• Подробную спецификацию\n"
        summary += "• Расчёт стоимости материалов\n"
        summary += "• Стоимость резов\n"
        summary += "• Вес изделий\n"
        summary += "• НДС (22%)\n"
        summary += "• Условия оплаты\n\n"
        summary += f"💰 Скидка: {discount_percent}%\n"
        summary += f"👤 Менеджер: {manager_name}\n"
        summary += f"📊 XLSX файл содержит расчётные формулы Excel!"
        
        # Сохраняем данные для генерации по нажатию кнопок
        await state.update_data(
            kp_order_data=order_data,
            kp_breakdown_tables=breakdown_tables,
            kp_offer_number=offer_number,
            kp_offer_date=offer_date,
            kp_db_id=kp_db_id,
            kp_pdf_path=None,
            kp_xlsx_path=None,
            kp_breakdown_path=None,
            kp_schema_pdf_path=None,
            kp_schema_breakdown_path=None,
            kp_customer_name=client_name,
            kp_manager_name=manager_name,
            kp_manager_phone=manager_phone,
            kp_manager_email=manager_email,
            kp_discount_percent=discount_percent,
            kp_transport_hours=transport_hours,
            kp_transport_price_per_hour=transport_price_per_hour,
            kp_delivery_conditions=delivery_conditions,
            kp_payment_conditions=payment_conditions,
        )
        
        # Очищаем состояние, чтобы callback мог сработать
        # Но данные остаются в state.data
        await state.set_state(None)
        
        logger.debug("Данные сохранены в state")
        logger.debug(f"  - Клиент: {client_name}")
        logger.debug(f"  - Менеджер: {manager_name}")
        logger.debug(f"  - Скидка: {discount_percent}%")
        logger.debug(f"  - Плит: {len(order_data)}")
        
        # Кнопки для генерации документов по нажатию (все доступны)
        await message.answer(
            f"{summary}\n\n"
            "💾 Хотите сохранить это КП в базу данных?\n\n"
            "Если сохраните, вы сможете отслеживать статус выполнения заказа.\n\n"
            "📥 Сгенерируйте и скачайте нужные файлы по кнопкам ниже:",
            reply_markup=save_to_db_with_files_kb(
                has_pdf=True,
                has_xlsx=True,
                has_breakdown=bool(breakdown_tables),
                has_schema=True,
                has_schema_breakdown=True,
            )
        )
    
    except FileGenerationError as e:
        # Ошибка генерации файлов
        logger.error(f"Ошибка генерации файлов КП для пользователя {message.from_user.id}: {e}", exc_info=True)
        await message.answer(
            f"❌ Не удалось создать файлы КП:\n{e}\n\n"
            "Попробуйте позже или обратитесь к администратору.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
    except Exception as e:
        # Непредвиденная ошибка
        logger.error(f"Непредвиденная ошибка при создании КП для пользователя {message.from_user.id}: {e}", exc_info=True)
        await message.answer(
            f"❌ Произошла ошибка: {str(e)}\n\n"
            "Проверьте формат данных и попробуйте снова.",
            reply_markup=main_menu_kb()
        )
        await state.clear()


# === СТАРЫЙ ОБРАБОТЧИК (для обратной совместимости) ===

@router.message(KPStates.waiting_for_commercial_offer)
async def receive_order_and_generate_pdf(message: Message, state: FSMContext):
    """Обработчик получения заказа и генерации PDF (старый обработчик)"""
    await message.answer("⏳ Формирую коммерческое предложение...")
    
    try:
        # Парсим список пользователя
        user_text = message.text or ""
        try:
            unparsed_lines, _line_contributions, _line_plate_load_details = set_plate_lists_from_text(
                user_text
            )
        except PlateParseError as e:
            # Ошибка парсинга - показываем понятное сообщение
            logger.warning(f"Ошибка парсинга заказа от пользователя {message.from_user.id}: {e}")
            await message.answer(
                f"❌ Не удалось распознать заказ:\n{e}\n\n"
                f"💡 Проверьте формат:\n"
                f"• ПБ 78-12-8п 5 шт\n"
                f"• 1.2x3.39 - 2\n"
                f"• 0,32x6,63 - 4",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return

        order = get_current_plate_order()
        await state.update_data(plate_order=order.to_dict())

        if unparsed_lines:
            # Показываем только первые 5 нераспознанных строк
            warn_text = "⚠️ Некоторые строки не распознаны:\n"
            warn_text += "\n".join(f"• {line}" for line in unparsed_lines[:5])
            if len(unparsed_lines) > 5:
                warn_text += f"\n... и ещё {len(unparsed_lines) - 5} строк"
            warn_text += (
                "\n\n💡 Я понимаю такие форматы:\n"
                "• ПБ 78-12-8п 3 шт\n"
                "• 1.2×3.39 — 2\n"
                "• 0,32x6,63 - 4\n"
                "• ПБ 66,2-12-8п 6"
            )
            await message.answer(warn_text)
        
        # ✅ ЗАПУСКАЕМ ОПТИМИЗАЦИЮ
        await message.answer("🔄 Оптимизирую раскрой плит для минимизации стоимости...")
        
        # Группируем плиты по армированию (из БД)
        orders_by_reinforcement = defaultdict(list)
        db_path = PB_DB_PATH
        
        if order.plate_load_details:
            logger.info("[COMMERCIAL] Используем PlateOrder (с нагрузками)")
            for key, qty in order.plate_load_details.items():
                length, width_m, load_code = key[0], key[1], key[2]
                ldr = key[3] if len(key) > 3 else order.plate_length_dm_raw.get(key, '')
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
                    'reinforcement': reinforcement_value,
                    'length_dm_raw': ldr,
                })
        
        # Запускаем оптимизацию для каждой группы армирования
        optimization_results_by_reinforcement = {}
        
        if orders_by_reinforcement:
            # Безопасная сортировка ключей
            keys_list = list(orders_by_reinforcement.keys())
            numeric_keys = sorted([k for k in keys_list if isinstance(k, (int, float))])
            string_keys = sorted([k for k in keys_list if isinstance(k, str)])
            all_keys_sorted = numeric_keys + string_keys
            
            logger.info(f"[COMMERCIAL] Найдено {len(orders_by_reinforcement)} групп(ы) по армированию")
            
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
                        
                        logger.info(
                            f"[COMMERCIAL] Армирование {reinforcement_key}: {optimization_result['total_plates']} плит"
                        )
                except Exception as e:
                    logger.exception(f"[COMMERCIAL] Ошибка оптимизации для армирования {reinforcement_key}: {e}")
            
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
                # Кешируем для отложенной генерации схемы
                OPT_PLAN_CACHE[message.from_user.id] = {
                    'plan': None,
                    'by_load': dict(optimization.OPT_CASCADING_PLAN_BY_LOAD),
                    'load_map': dict(optimization.LOAD_TO_REINFORCEMENT_MAP),
                }
                logger.info(
                    f"[COMMERCIAL] Сохранено {len(optimization_results_by_reinforcement)} результатов оптимизации"
                )
                
                await message.answer("✅ Оптимизация завершена! Формирую документы...")
        
        # 🔥 ТЕПЕРЬ build_price_rows получит ОПТИМИЗИРОВАННЫЕ данные из OPT_CASCADING_PLAN_BY_LOAD!
        from viz_modules.procurement import build_price_rows, build_component_breakdown, build_procurement_items, get_orders_from_opt_plan
        from viz_modules.price_utils import load_price_table_from_xlsx
        
        # Загружаем таблицу цен для расчётов
        price_table = load_price_table_from_xlsx(str(cfg.PRICE_XLSX_PATH))
        
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
        
        # Формируем order_data из того же источника, что и смета (build_procurement_items)
        order_data = []
        procurement_items = build_procurement_items()
        for item in procurement_items:
            length_m = item['length']
            width_m = item['width']
            qty = item['qty']
            load_code = item.get('load_code')
            if load_code is None:
                load_code = cfg.get_load_code_for_plate(length_m, width_m, default=(6 if width_m < 1.0 else 8))
            matching_row = None
            for row in price_rows:
                if len(row) < 8:
                    continue
                row_name = row[1]
                row_qty = row[2]
                parsed_length, parsed_width = cfg.parse_name_to_sizes(row_name)
                if parsed_length is None or parsed_width is None:
                    continue
                parsed_load = cfg.parse_load_code_from_name(row_name)
                # Сопоставление по длине с допуском 0.01 м: различаем 6.35 и 6.37 (не склеиваем в один round).
                # Для 12,5п: из имени парсится 13, в заказе 12.5 — сравниваем через load_code_for_price_match
                length_match = abs(parsed_length - length_m) < 0.01
                if (length_match
                        and abs(parsed_width - width_m) < 0.01
                        and cfg.load_code_for_price_match(parsed_load) == cfg.load_code_for_price_match(load_code)):
                    matching_row = row
                    break
            if matching_row:
                name = matching_row[1]
                price_str = matching_row[7]
                try:
                    unit_price = float(price_str.replace(' ', '').replace(',', '.'))
                except (ValueError, AttributeError):
                    unit_price = 0.0
                match = re.search(r'ПБ\s+([\d,]+)-', name)
                length_dm_raw = match.group(1).strip() if match else ''
                # Не подменять имя/длину, если у позиции явно задан length_dm_raw и он отличается от подстроки в строке сметы (57 vs 57,1)
                item_ldr = item.get('length_dm_raw')
                if item_ldr and length_dm_raw and length_dm_raw != item_ldr:
                    name = cfg.make_plate_name(length_m, width_m, load_code=load_code, length_dm_raw=item_ldr)
                    length_dm_raw = item_ldr
            else:
                # #region agent log
                try:
                    _log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'debug-b59370.log')
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "b59370", "hypothesisId": "H3", "location": "commercial:order_data_no_match", "message": "no matching_row, using make_plate_name", "data": {"length_m": length_m, "width_m": width_m}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
                # #endregion
                # length_dm_raw из item прокинут из build_procurement_items (57,1 не теряется)
                item_ldr = item.get('length_dm_raw') or None
                name = cfg.make_plate_name(length_m, width_m, load_code=load_code, length_dm_raw=item_ldr)
                unit_price = 0.0
                length_dm_raw = item_ldr or (
                    f'{length_m * 10:.1f}'.rstrip('0').rstrip('.').replace('.', ',') if length_m else ''
                )
            # Если кэш номенклатуры уже содержит canonical_name для этой позиции —
            # используем его вместо имени из сметы/make_plate_name, а nomenclature_id
            # подставляем сразу, чтобы enrich не делал повторный поиск по БД.
            cache_name = item.get('canonical_name')
            cache_nid = item.get('nomenclature_id')
            if cache_name:
                name = cache_name
                _raw_match = re.search(r'ПБ\s+([\d,]+)-', name)
                length_dm_raw = _raw_match.group(1).strip() if _raw_match else length_dm_raw
            # Вес КП: plate_weights по размерам (сумма на строку — для save_kp_to_db)
            _, total_weight_kg = resolve_kp_line_weight_kg({
                "length_m": length_m,
                "width_m": width_m,
                "qty": qty,
            })
            # #region agent log (57/57,1: итоговое имя в order_data, альт. поток)
            if 5.69 <= length_m <= 5.73:
                try:
                    _log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'debug-8e9428.log')
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_order_data_alt", "location": "commercial:order_data_loop_alt", "message": "57/57,1: final name in order_data", "data": {"length_m": length_m, "width_m": width_m, "qty": qty, "name": name, "from_matching_row": matching_row[1] if matching_row else None}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            # #region agent log (a9176e: 57/57,1 — order_data alt flow)
            if 5.69 <= length_m <= 5.73:
                try:
                    _parsed = (cfg.parse_name_to_sizes(matching_row[1]) if matching_row and len(matching_row) > 1 else (None, None))
                    _log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'debug-a9176e.log')
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "a9176e", "hypothesisId": "H3", "location": "commercial:order_data_loop_alt", "message": "57/57,1 order_data name source", "data": {"length_m": length_m, "qty": qty, "has_matching_row": matching_row is not None, "matching_row_name": matching_row[1] if matching_row and len(matching_row) > 1 else None, "parsed_length": _parsed[0], "item_length_dm_raw": item.get('length_dm_raw'), "cache_name": item.get('canonical_name'), "final_name": name, "final_length_dm_raw": length_dm_raw}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            entry = {
                "name": name,
                "length_m": length_m,
                "length_dm_raw": length_dm_raw,
                "width_m": width_m,
                "qty": qty,
                "load_class": (cfg.normalize_load_code(load_code) or 8) * 100 if load_code is not None else 800,
                "unit_price": unit_price,
                "weight": total_weight_kg,
            }
            if cache_nid is not None:
                entry['nomenclature_id'] = cache_nid
            order_data.append(entry)
        total_order_data = sum(i['qty'] for i in order_data)
        plan_orders = get_orders_from_opt_plan() or []
        total_plan = sum(o['qty'] for o in plan_orders)
        if total_plan and total_order_data != total_plan:
            logger.error(f"[КП] ПОТЕРЯ ПЛИТ: order_data={total_order_data}, plan={total_plan}")
            await message.answer(
                "⚠️ ВНИМАНИЕ: Обнаружено расхождение в количестве плит!\n"
                f"В плане: {total_plan} шт\n"
                f"В смете: {total_order_data} шт\n\n"
                "Пожалуйста, проверьте данные и попробуйте заново.",
                reply_markup=main_menu_kb()
            )
            await state.clear()
            return
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
        
        # Документы генерируются по нажатию кнопок — показываем кнопки сразу
        # Формируем сводку по заказу
        total_qty = sum(item['qty'] for item in order_data)
        summary = f"✅ Коммерческое предложение готово!\n\n"
        summary += f"📋 Заказ:\n"
        for item in order_data:
            summary += f"  • {item['name']} — {item['qty']} шт\n"
        summary += f"\n📊 Всего позиций: {len(order_data)}\n"
        summary += f"📦 Всего плит: {total_qty} шт\n\n"
        summary += "✨ Документы содержат:\n"
        summary += "• Подробную спецификацию\n"
        summary += "• Расчёт стоимости материалов\n"
        summary += "• Стоимость резов\n"
        summary += "• Вес изделий\n"
        summary += "• НДС (22%)\n"
        summary += "• Условия оплаты\n\n"
        summary += "📊 XLSX файл содержит расчётные формулы Excel!\n"
        summary += "📐 Схема раскладки и детальная разбивка — по кнопкам."
        
        # Сохраняем данные для генерации по нажатию кнопок
        await state.update_data(
            kp_order_data=order_data,
            kp_breakdown_tables=breakdown_tables,
            kp_offer_number=offer_number,
            kp_offer_date=offer_date,
            kp_db_id=None,
            kp_pdf_path=None,
            kp_xlsx_path=None,
            kp_breakdown_path=None,
            kp_schema_pdf_path=None,
            kp_schema_breakdown_path=None,
            kp_customer_name=customer_name,
            kp_manager_name=None,
            kp_manager_phone=None,
            kp_manager_email=None,
            kp_discount_percent=0,
            kp_delivery_conditions=None,
            kp_payment_conditions=None,
        )
        
        await state.set_state(None)
        
        # Кнопки для генерации документов по нажатию
        await message.answer(
            f"{summary}\n\n"
            "💾 Хотите сохранить это КП в базу данных?\n\n"
            "Если сохраните, вы сможете отслеживать статус выполнения заказа.\n\n"
            "📥 Сгенерируйте и скачайте нужные файлы по кнопкам ниже:",
            reply_markup=save_to_db_with_files_kb(
                has_pdf=True,
                has_xlsx=True,
                has_breakdown=bool(breakdown_tables),
                has_schema=True,
                has_schema_breakdown=True,
            )
        )
    
    except FileGenerationError as e:
        # Ошибка генерации файлов
        logger.error(f"Ошибка генерации файлов КП (старый обработчик) для пользователя {message.from_user.id}: {e}", exc_info=True)
        await message.answer(
            f"❌ Не удалось создать файлы КП:\n{e}\n\n"
            "Попробуйте позже или обратитесь к администратору.",
            reply_markup=main_menu_kb()
        )
        await state.clear()
    except Exception as e:
        # Непредвиденная ошибка
        logger.error(f"Непредвиденная ошибка при создании КП (старый обработчик) для пользователя {message.from_user.id}: {e}", exc_info=True)
        await message.answer(
            f"❌ Произошла ошибка: {str(e)}\n\n"
            "Проверьте формат данных и попробуйте снова.",
            reply_markup=main_menu_kb()
        )
        await state.clear()


# ==================== ОБРАБОТЧИКИ СКАЧИВАНИЯ ФАЙЛОВ КП ====================

FILE_TYPE_TO_KEY = {
    "pdf": "kp_pdf_path",
    "xlsx": "kp_xlsx_path",
    "breakdown": "kp_breakdown_path",
    "schema": "kp_schema_pdf_path",
    "schema_breakdown": "kp_schema_breakdown_path",
}

FILE_TYPE_CAPTIONS = {
    "pdf": "📄 Коммерческое предложение (PDF)",
    "xlsx": "📊 Коммерческое предложение (XLSX)",
    "breakdown": "📋 Детальная разбивка компонентов",
    "schema": "📐 Схема раскладки плит",
    "schema_breakdown": "📊 Детальная разбивка по схеме",
}


@router.callback_query(F.data.startswith("kp_file_"))
async def callback_kp_file_download(callback: CallbackQuery, state: FSMContext):
    """Генерирует и отправляет файл КП по нажатию кнопки. Если файл ещё не создан — создаёт."""
    await callback.answer()
    
    file_type = callback.data.replace("kp_file_", "")
    if file_type not in FILE_TYPE_TO_KEY:
        await callback.message.answer("❌ Неизвестный тип файла.")
        return
    
    data = await state.get_data()
    path = data.get(FILE_TYPE_TO_KEY[file_type])
    
    # Если файл уже есть — отправляем
    if path and os.path.exists(path):
        try:
            await callback.message.answer_document(
                FSInputFile(path),
                caption=FILE_TYPE_CAPTIONS.get(file_type, "Файл")
            )
        except Exception as e:
            logger.exception(f"Ошибка отправки файла {file_type}: {e}")
            await callback.message.answer(f"❌ Не удалось отправить файл: {e}")
        return
    
    # Генерируем по требованию
    order_data = data.get('kp_order_data', [])
    if not order_data:
        await callback.message.answer("❌ Данные КП недоступны. Создайте КП заново.")
        return
    
    offer_number = data.get('kp_offer_number', f"kp_{datetime.now().strftime('%Y%m%d%H%M')}")
    offer_date = data.get('kp_offer_date', datetime.now().strftime("%d.%m.%Y"))
    client_name = data.get('kp_customer_name', 'Клиент')
    manager_name = data.get('kp_manager_name') or ''
    manager_phone = data.get('kp_manager_phone') or ''
    manager_email = data.get('kp_manager_email') or ''
    discount_percent = data.get('kp_discount_percent', 0)
    delivery_conditions = data.get('kp_delivery_conditions') or ''
    payment_conditions = data.get('kp_payment_conditions') or ''
    kp_db_id = data.get('kp_db_id')
    
    try:
        if file_type == "pdf":
            await callback.message.answer("⏳ Генерирую PDF...")
            pdf_buffer = await asyncio.to_thread(
                generate_commercial_offer_pdf,
                order_data, offer_number, offer_date,
                client_name, manager_name, manager_phone, manager_email,
                discount_percent, kp_db_id
            )
            path = os.path.join(OUTPUTS_DIR_STR, f"КП_{offer_number}_{offer_date.replace('.', '')}.pdf")
            with open(path, 'wb') as f:
                f.write(pdf_buffer.getvalue())
            await state.update_data(kp_pdf_path=path)
            await callback.message.answer_document(FSInputFile(path), caption=FILE_TYPE_CAPTIONS["pdf"])
            
        elif file_type == "xlsx":
            await callback.message.answer("⏳ Генерирую XLSX...")
            xlsx_buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data, offer_number, offer_date,
                client_name, manager_name, manager_phone, manager_email,
                discount_percent, delivery_conditions, payment_conditions, kp_db_id
            )
            path = os.path.join(OUTPUTS_DIR_STR, f"КП_{offer_number}_{offer_date.replace('.', '')}.xlsx")
            with open(path, 'wb') as f:
                f.write(xlsx_buffer.getvalue())
            await state.update_data(kp_xlsx_path=path)
            await callback.message.answer_document(FSInputFile(path), caption=FILE_TYPE_CAPTIONS["xlsx"])
            
        elif file_type == "breakdown":
            breakdown_tables = data.get('kp_breakdown_tables', [])
            if not breakdown_tables:
                await callback.message.answer("❌ Детальная разбивка недоступна.")
                return
            await callback.message.answer("⏳ Генерирую разбивку...")
            from core.commercial_offer import save_breakdown_to_excel
            path = os.path.join(OUTPUTS_DIR_STR, f"Детальная_разбивка_{offer_number}_{offer_date.replace('.', '')}.xlsx")
            await asyncio.to_thread(save_breakdown_to_excel, breakdown_tables, path)
            if os.path.exists(path):
                await state.update_data(kp_breakdown_path=path)
                await callback.message.answer_document(FSInputFile(path), caption=FILE_TYPE_CAPTIONS["breakdown"])
            else:
                await callback.message.answer("❌ Не удалось создать файл разбивки.")
                
        elif file_type in ("schema", "schema_breakdown"):
            # Восстанавливаем план оптимизации из кеша (для корректной схемы)
            user_id = callback.from_user.id
            if user_id in OPT_PLAN_CACHE:
                import core.optimization as optimization
                cached = OPT_PLAN_CACHE[user_id]
                optimization.OPT_CASCADING_PLAN = cached.get('plan')
                optimization.OPT_CASCADING_PLAN_BY_LOAD = cached.get('by_load', {})
                optimization.LOAD_TO_REINFORCEMENT_MAP = cached.get('load_map', {})
            
            schema_pdf_path = data.get('kp_schema_pdf_path')
            schema_breakdown_path = data.get('kp_schema_breakdown_path')
            
            if (not schema_pdf_path or not os.path.exists(schema_pdf_path)) or \
               (file_type == "schema_breakdown" and (not schema_breakdown_path or not os.path.exists(schema_breakdown_path))):
                await callback.message.answer("⏳ Генерирую схему раскладки...")
                result_paths = await asyncio.to_thread(visualize_plan, OUTPUTS_DIR_STR)
                if isinstance(result_paths, tuple) and len(result_paths) >= 2:
                    png_path, pdf_schema_path = result_paths
                    base = os.path.basename(png_path)
                    timestamp = base.split('КЗ_', 1)[-1].replace('.png', '') if 'КЗ_' in base else base.rsplit('_', 1)[-1].replace('.png', '')
                    schema_breakdown_path = os.path.join(OUTPUTS_DIR_STR, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx')
                    schema_pdf_path = pdf_schema_path if os.path.exists(pdf_schema_path) else None
                    await state.update_data(kp_schema_pdf_path=schema_pdf_path, kp_schema_breakdown_path=schema_breakdown_path)
            
            send_path = schema_pdf_path if file_type == "schema" else schema_breakdown_path
            if send_path and os.path.exists(send_path):
                await callback.message.answer_document(FSInputFile(send_path), caption=FILE_TYPE_CAPTIONS[file_type])
            else:
                await callback.message.answer("❌ Не удалось сгенерировать схему.")
                
    except Exception as e:
        logger.exception(f"Ошибка генерации {file_type}: {e}")
        await callback.message.answer(f"❌ Ошибка при создании файла: {e}")


# ==================== ОБРАБОТЧИКИ СОХРАНЕНИЯ КП В БД ====================

@router.callback_query(F.data == "save_kp_to_db")
async def callback_save_kp_to_db(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "Сохранить в БД".
    Запрашивает сроки выполнения КП.
    """
    logger.debug("Нажата кнопка 'Сохранить в БД'")
    
    # Убираем "часики" с кнопки
    await callback.answer()
    
    # === РАСЧЕТ КОЛИЧЕСТВА ДОРОЖЕК И СРОКОВ ===
    # Получаем данные КП из состояния
    data = await state.get_data()
    order_data = data.get('kp_order_data', [])
    
    # Подсчитываем количество дорожек для планирования производства
    MAX_TRACK_LENGTH = 101.0
    total_length = 0.0
    
    # Суммируем длины всех плит с учетом количества
    for item in order_data:
        length = item.get('length_m', 0)
        qty = item.get('qty', 1)
        total_length += length * qty
    
    # Простой расчет: делим общую длину на максимальную длину дорожки
    # (это приблизительная оценка, точный расчет делается при планировании производства)
    estimated_tracks = max(1, int(round(total_length / MAX_TRACK_LENGTH + 0.5)))
    
    # Расчет сроков: количество дорожек / 5 (производительность 5 дорожек в день)
    estimated_days = max(1, int(round(estimated_tracks / 5.0 + 0.5)))
    
    logger.debug(f"Расчет сроков: общая длина={total_length:.1f}м, дорожек≈{estimated_tracks}, дней≈{estimated_days}")
    
    # Редактируем сообщение с кнопками
    await callback.message.edit_text(
        "✅ Сохраняю КП в базу данных...\n\n"
        f"⏱️ Оценка производства:\n"
        f"  • Примерно дорожек: {estimated_tracks}\n"
        f"  • Примерно дней: {estimated_days}\n\n"
        "📅 Укажите сроки выполнения:\n"
        f"(Например: '{estimated_days} дней', '2 недели', '01.02.2024')"
    )
    
    logger.debug("Переход к состоянию waiting_execution_terms")
    
    # Переходим к состоянию ожидания сроков
    await state.set_state(KPStates.waiting_execution_terms)


@router.message(KPStates.waiting_execution_terms)
async def receive_execution_terms(message: Message, state: FSMContext):
    """
    Обработчик ввода сроков выполнения.
    Сохраняет КП в базу данных plita.db.
    """
    execution_terms_input = message.text.strip()
    logger.debug(f"Получены сроки: {execution_terms_input}")
    
    # === ПАРСИМ СРОКИ И ВЫЧИСЛЯЕМ ДАТУ ДЕДЛАЙНА ===
    from datetime import timedelta
    import re
    
    deadline_date = None
    
    # ИСПРАВЛЕНИЕ: Сначала пробуем распознать ДАТУ (чтобы "01.02.2026" не распознавалось как "1 день")
    # Вариант 1: Формат ДД.ММ.ГГГГ (например: "01.02.2026")
    try:
        deadline_date = datetime.strptime(execution_terms_input, '%d.%m.%Y')
        logger.debug(f"Распознана дата (ДД.ММ.ГГГГ): {deadline_date.strftime('%d.%m.%Y')}")
    except ValueError:
        pass
    
    # Вариант 2: Формат ГГГГ-ММ-ДД (например: "2026-02-01")
    if not deadline_date:
        try:
            deadline_date = datetime.strptime(execution_terms_input, '%Y-%m-%d')
            logger.debug(f"Распознана дата (ГГГГ-ММ-ДД): {deadline_date.strftime('%d.%m.%Y')}")
        except ValueError:
            pass
    
    # Вариант 3: Пользователь ввёл количество дней (например: "14", "14 дней", "30дней")
    if not deadline_date:
        match_days = re.search(r'(\d+)\s*(?:дн|день|дней|day|days)', execution_terms_input, re.IGNORECASE)
        if match_days:
            days = int(match_days.group(1))
            deadline_date = datetime.now() + timedelta(days=days)
            logger.debug(f"Распознано {days} дней, дедлайн: {deadline_date.strftime('%d.%m.%Y')}")
    
    # Вариант 4: Пользователь ввёл количество недель (например: "2 недели", "3week")
    if not deadline_date:
        match_weeks = re.search(r'(\d+)\s*(?:нед|недел|недели|week|weeks)', execution_terms_input, re.IGNORECASE)
        if match_weeks:
            weeks = int(match_weeks.group(1))
            deadline_date = datetime.now() + timedelta(weeks=weeks)
            logger.debug(f"Распознано {weeks} недель, дедлайн: {deadline_date.strftime('%d.%m.%Y')}")
    
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
    logger.debug(f"Итоговая дата дедлайна: {execution_terms}")
    
    # Получаем данные КП из состояния
    data = await state.get_data()
    order_data = data.get('kp_order_data', [])
    xlsx_path = data.get('kp_xlsx_path')
    customer_name = data.get('kp_customer_name')
    manager_name = data.get('kp_manager_name')
    manager_phone = data.get('kp_manager_phone', '')
    manager_email = data.get('kp_manager_email', '')
    discount_percent = data.get('kp_discount_percent', 0)
    transport_hours = data.get('kp_transport_hours')
    transport_price_per_hour = data.get('kp_transport_price_per_hour')
    delivery_conditions = data.get('kp_delivery_conditions')
    payment_conditions = data.get('kp_payment_conditions')
    offer_date = data.get('kp_offer_date')
    offer_number = data.get('kp_offer_number', '')
    kp_db_id = data.get('kp_db_id')
    
    # Генерируем XLSX, если ещё не создан
    if not xlsx_path or not os.path.exists(xlsx_path):
        try:
            await message.answer("⏳ Генерирую XLSX для сохранения в БД...")
            xlsx_buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data, offer_number, offer_date,
                customer_name, manager_name, manager_phone, manager_email,
                discount_percent, delivery_conditions, payment_conditions, kp_db_id
            )
            xlsx_path = os.path.join(OUTPUTS_DIR_STR, f"КП_{offer_number}_{offer_date.replace('.', '')}.xlsx")
            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_buffer.getvalue())
            await state.update_data(kp_xlsx_path=xlsx_path)
        except Exception as e:
            logger.exception(f"Ошибка генерации XLSX для save: {e}")
    
    logger.debug("Данные из state")
    logger.debug(f"  - Клиент: {customer_name}")
    logger.debug(f"  - Менеджер: {manager_name}")
    logger.debug(f"  - Скидка: {discount_percent}%")
    logger.debug(f"  - Плит в заказе: {len(order_data)}")
    logger.debug(f"  - XLSX путь: {xlsx_path}")
    logger.debug(f"  - Дата КП: {offer_date}")
    
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
            status='в работе',
            transport_hours=transport_hours,
            transport_price_per_hour=transport_price_per_hour,
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
        logger.exception(f"Ошибка при сохранении КП в БД: {e}")
        await message.answer(
            "❌ Не удалось сохранить КП в базу данных.\n\n"
            "Попробуйте позже. Если ошибка повторяется — смотри logs/bot.log.",
            reply_markup=main_menu_kb()
        )
    
    finally:
        user_id = message.from_user.id
        if user_id in OPT_PLAN_CACHE:
            del OPT_PLAN_CACHE[user_id]
        await state.clear()


@router.callback_query(F.data == "skip_save_kp")
async def callback_skip_save_kp(callback: CallbackQuery, state: FSMContext):
    """
    Обработчик кнопки "Не сохранять".
    Пропускает сохранение КП в БД.
    """
    await callback.answer()
    user_id = callback.from_user.id
    if user_id in OPT_PLAN_CACHE:
        del OPT_PLAN_CACHE[user_id]
    
    await callback.message.edit_text("❌ КП не сохранено в базу данных.")
    await callback.message.answer("✅ Работа с КП завершена!", reply_markup=main_menu_kb())
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
    logger.debug("Нажата кнопка 'В архив'")
    
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
    manager_phone = data.get('kp_manager_phone', '')
    manager_email = data.get('kp_manager_email', '')
    discount_percent = data.get('kp_discount_percent', 0)
    transport_hours = data.get('kp_transport_hours')
    transport_price_per_hour = data.get('kp_transport_price_per_hour')
    delivery_conditions = data.get('kp_delivery_conditions')
    payment_conditions = data.get('kp_payment_conditions')
    offer_date = data.get('kp_offer_date')
    offer_number = data.get('kp_offer_number', '')
    kp_db_id = data.get('kp_db_id')
    
    # Генерируем XLSX, если ещё не создан
    if (not xlsx_path or not os.path.exists(xlsx_path)) and order_data:
        try:
            await callback.message.answer("⏳ Генерирую XLSX для сохранения в архив...")
            xlsx_buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data, offer_number, offer_date,
                customer_name, manager_name, manager_phone, manager_email,
                discount_percent, delivery_conditions, payment_conditions, kp_db_id
            )
            xlsx_path = os.path.join(OUTPUTS_DIR_STR, f"КП_{offer_number}_{offer_date.replace('.', '')}.xlsx")
            with open(xlsx_path, 'wb') as f:
                f.write(xlsx_buffer.getvalue())
            await state.update_data(kp_xlsx_path=xlsx_path)
        except Exception as e:
            logger.exception(f"Ошибка генерации XLSX для archive: {e}")
    
    logger.debug("Данные из state")
    logger.debug(f"  - Клиент: {customer_name}")
    logger.debug(f"  - Менеджер: {manager_name}")
    logger.debug(f"  - Скидка: {discount_percent}%")
    logger.debug(f"  - Плит в заказе: {len(order_data)}")
    logger.debug(f"  - XLSX путь: {xlsx_path}")
    logger.debug(f"  - Дата КП: {offer_date}")
    
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
            status='в архиве',  # СТАТУС АРХИВА!
            transport_hours=transport_hours,
            transport_price_per_hour=transport_price_per_hour,
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
        logger.exception(f"Ошибка при сохранении КП в архив: {e}")
        await callback.message.answer(
            "❌ Не удалось сохранить КП в архив.\n\n"
            "Попробуйте позже. Если ошибка повторяется — смотри logs/bot.log.",
            reply_markup=main_menu_kb()
        )
    
    finally:
        user_id = callback.from_user.id
        if user_id in OPT_PLAN_CACHE:
            del OPT_PLAN_CACHE[user_id]
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

