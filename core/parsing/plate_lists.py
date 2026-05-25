# -*- coding: utf-8 -*-
"""Парсинг текста заказа в списки плит (мутабельный рантайм)."""

from __future__ import annotations

import logging
import re
from typing import Any, Optional, Tuple

from ..domain.plate_order import _try_fill_plate_nomenclature_cache
from ..exceptions import PlateParseError
from ..plate_line_parser import parse_line
from ..plate_runtime_state import get_plate_mutable_runtime
from ..plate_validation import validate_plate_values
from ..runtime import NomenclatureCacheFiller

logger = logging.getLogger(__name__)

# Ключ для сопоставления строки ввода с позициями заказа / order_data:
# длина и ширина в метрах, код нагрузки (или None для размерных строк без марки), сырой фрагмент длины из марки.
LineContributionKey = Tuple[float, float, Optional[float], str]


def get_last_parse_diagnostics() -> list[dict[str, Any]]:
    """Возвращает диагностику последнего запуска set_plate_lists_from_text()."""
    return list(get_plate_mutable_runtime().last_parse_diagnostics)


def _clear_all_plate_lists() -> None:
    """Очищает списки плит в текущем рантайме заказа."""
    get_plate_mutable_runtime().clear_plate_lists()


def _recompute_totals_from_lists() -> None:
    """Пересчитывает итоговые поля рантайма на основе списков плит."""
    get_plate_mutable_runtime()._recompute_totals_from_lists()


def set_plate_lists_from_text(
    user_text: str,
    *,
    fill_nomenclature_cache: NomenclatureCacheFiller | None = None,
) -> tuple[list[str], list[list[LineContributionKey]], list[dict[tuple, int]]]:
    """Парсит свободный текст пользователя и заполняет списки PLATES_*.

    Поддерживаемые форматы (регистр не важен, пробелы опциональны):
      - Размеры через «x» или «×»: «1.2×3.39 — 2 шт», «0,32x6,63 - 4»
      - Марка ПБ/ПК: «ПБ 78-12-8п 3», «Плиты ПБ 78-12-8п», «ПБ78-12-8п 5», «ПК 80-12-8 7»
        Длина и ширина в дециметрах (78 => 7.8 м, 12 => 1.2 м), нагрузка — после последнего дефиса (8п, 10п и т.д.)
      - Марка без префикса ПБ: «71-12-8 3», «65,6-12-12,5 2», «71-9,20-8 1»
      - Количество: после марки, опционально «шт» («8п 5», «8п 5 шт», «8п — 5»)
    Неизвестные ширины и нераспознанные строки возвращаются в списке нераспознанных.

    Returns:
        (unparsed_lines, line_contributions, line_plate_load_details): нераспознанные строки;
        для каждой строки ``lines`` — список ключей вклада; для каждой строки — словарь
        накопленных количеств по тем же ключам, что ``PLATE_LOAD_DETAILS``, только для этой строки ввода.

    Raises:
        PlateParseError: Если текст пустой или после разбивки не осталось валидных строк.
    """
    if not user_text or not user_text.strip():
        logger.warning("Получен пустой текст заказа")
        raise PlateParseError(
            "Текст заказа пустой. Пожалуйста, введите список плит.\n"
            "Пример: ПБ 78-12-8п 5 шт"
        )

    _clear_all_plate_lists()
    rt = get_plate_mutable_runtime()
    rt.last_parse_diagnostics.clear()

    # Нормализация: конвертация каталожных марок (ПБ 59.12-8Вр1400-25 → ПБ 59-12-8п)
    # и других нестандартных вариантов записи перед основным парсингом.
    _processing_text = user_text
    try:
        from ..plate_text_normalizer import normalize_order_text

        _norm = normalize_order_text(user_text)
        if _norm.warnings:
            for _w in _norm.warnings[:10]:
                logger.info("Нормализатор: %s", _w)
        if _norm.normalized_text.strip():
            _processing_text = _norm.normalized_text
    except Exception as _norm_err:
        logger.warning("Ошибка нормализатора, используем исходный текст: %s", _norm_err)

    # Нормализация: единый символ умножения, неразрывные пробелы как обычные
    text = (_processing_text or "").replace("\u00d7", "x").replace("×", "x")
    text = text.replace("\u00a0", " ")
    lines = [re.sub(r"\s+", " ", l).strip() for l in re.split(r"[\n;]+", text) if l.strip()]
    line_contributions: list[list[LineContributionKey]] = [[] for _ in lines]
    line_plate_load_details: list[dict[tuple, int]] = [{} for _ in lines]

    def _record_contribution(line_idx: int, length_m: float, width_m: float, load_code: Optional[float], ldr: str) -> None:
        if line_idx < 0 or line_idx >= len(line_contributions):
            return
        line_contributions[line_idx].append(
            (round(float(length_m), 3), round(float(width_m), 3), load_code, (ldr or "").strip())
        )

    # Дополнительная проверка после разбивки на строки
    if not lines:
        logger.warning("После разбивки не осталось валидных строк")
        raise PlateParseError(
            "Не удалось найти ни одной строки с плитами.\n"
            "Проверьте формат ввода."
        )

    def add_items(
        width_m: float,
        length_m: float,
        qty: int,
        load_code: int = None,
        length_dm_raw: str = None,
        line_idx: Optional[int] = None,
    ):
        """
        Добавляет плиты в соответствующий глобальный список по ширине.

        Если ширина не попадает ни в один из «жёстких» диапазонов,
        используем правило: берём ближайший МЕНЬШИЙ допустимый рез
        (включая остаточные ширины) и повторно вызываем add_items
        уже с притянутой шириной.

        Args:
            width_m: ширина плиты в метрах
            length_m: длина плиты в метрах
            qty: количество плит
            load_code: нагрузка (6, 8, 10, 12, 13 и т.д.) - опционально
            length_dm_raw: исходная строка длины из марки (например "59,81") для различения плит
            line_idx: индекс строки заказа (для line_contributions); None — не писать вклад
        """
        # ЗАЩИТА: Проверяем адекватность размеров
        # Если размеры слишком большие (вероятно, ошибка OCR распознал мм как дм)
        # то игнорируем эту плиту
        if width_m > 20.0 or length_m > 200.0:
            logger.warning(
                f"Пропущена плита с неадекватными размерами: {length_m}м × {width_m}м. "
                f"Возможно, OCR распознал мм как дм."
            )
            return

        if width_m <= 0 or length_m <= 0:
            logger.warning(f"Пропущена плита с нулевыми размерами: {length_m}м × {width_m}м")
            return

        # Защита от зависаний: слишком большое количество может «повесить» бот,
        # потому что ниже мы добавляем в списки по 1 штуке в цикле.
        if qty is None or qty <= 0:
            logger.warning(f"Пропущена плита с некорректным количеством: qty={qty}")
            return
        if qty > 500:
            logger.warning(f"Пропущена строка с слишком большим количеством плит: qty={qty}")
            return
        # Специальная обработка плит 1.5 м → заменяем на 1.2 м + 0.3 м
        if 1.45 <= width_m <= 1.55:  # 1.5 м (диапазон ±50 мм)
            length_rounded = round(float(length_m), 3)
            # Добавляем плиту 1.2 м
            for _ in range(max(0, qty)):
                rt.plates_1_2.append(length_rounded)
                # Сохраняем точную ширину 1.2м
                rt.plate_exact_widths[(length_rounded, "PLATES_1_2")] = 1.2
            # Добавляем плиту 0.3 м (записываем в PLATES_0_32)
            for _ in range(max(0, qty)):
                rt.plates_0_32.append(length_rounded)
                # Сохраняем точную ширину 0.3м (попадает в диапазон 0.26-0.32)
                rt.plate_exact_widths[(length_rounded, "PLATES_0_32")] = 0.3
            ldr_norm = (length_dm_raw or "").strip()
            if load_code is not None and load_code > 0:
                width_rounded = round(width_m, 3)
                key_new = (length_rounded, width_rounded, load_code, ldr_norm)
                rt.plate_load_details[key_new] = rt.plate_load_details.get(key_new, 0) + qty
                rt.plate_length_dm_raw[key_new] = ldr_norm
                if line_idx is not None and 0 <= line_idx < len(line_plate_load_details):
                    _ld = line_plate_load_details[line_idx]
                    _ld[key_new] = _ld.get(key_new, 0) + qty
                if line_idx is not None:
                    lc = float(load_code)
                    _record_contribution(line_idx, length_rounded, 1.2, lc, ldr_norm)
                    _record_contribution(line_idx, length_rounded, 0.3, lc, ldr_norm)
            elif line_idx is not None:
                _record_contribution(line_idx, length_rounded, 1.2, None, ldr_norm)
                _record_contribution(line_idx, length_rounded, 0.3, None, ldr_norm)
            return

        target = None
        target_name = None  # Имя списка для сохранения точной ширины

        # Стандартные ширины плит
        if 1.15 <= width_m <= 1.25:
            target = rt.plates_1_2
            target_name = "PLATES_1_2"
        elif 0.98 <= width_m <= 1.02:
            target = rt.plates_1_0
            target_name = "PLATES_1_0"
        elif 1.02 <= width_m <= 1.08:  # по таблице завода: рез 1020–1080 мм
            target = rt.plates_1_08
            target_name = "PLATES_1_08"
        # Основные части (по таблице допустимых резов: 260-320, 460-530, 660-720, 860-920):
        elif 0.26 <= width_m <= 0.32:  # 260-320 мм
            target = rt.plates_0_32
            target_name = "PLATES_0_32"
        elif 0.46 <= width_m <= 0.53:  # 460-530 мм
            target = rt.plates_0_46
            target_name = "PLATES_0_46"
        elif 0.66 <= width_m <= 0.71:  # 660-710 мм → PLATES_0_70
            target = rt.plates_0_70
            target_name = "PLATES_0_70"
        elif 0.71 < width_m <= 0.72:  # 710-720 мм → PLATES_0_72
            target = rt.plates_0_72
            target_name = "PLATES_0_72"
        elif 0.86 <= width_m <= 0.92:  # 860-920 мм
            target = rt.plates_0_86
            target_name = "PLATES_0_86"
        # Остатки по таблице завода (добор): остаток от 860–920 = 260–320 (попадает в 0_32 выше)
        # 340 мм по таблице не входит в допустимый остаток — не выделяем отдельно
        elif 0.47 < width_m <= 0.49:  # ~480 мм (остаток от 720)
            target = rt.plates_0_48
            target_name = "PLATES_0_48"
        elif 0.49 < width_m <= 0.51:  # ~500 мм (остаток от 700)
            target = rt.plates_0_50
            target_name = "PLATES_0_50"
        # По таблице остаток от реза 460–530 = 660–720 мм (попадает в 0_70/0_72 выше), 740 не используем
        elif 0.87 < width_m <= 0.89:  # ~880 мм (остаток от 320)
            target = rt.plates_0_88
            target_name = "PLATES_0_88"
        else:
            # Здесь ширина не попала ни в один диапазон.
            # Применяем правило: «берём меньший рез».
            # Допустимые ширины по таблице завода (информ. письмо): резы 260–320, 460–530, 660–720, 860–920, 1020–1080 мм
            STANDARD_WIDTHS = [
                0.20,
                0.30,  # специальные ленты
                0.32,  # рез 300 (-40;+20) = 260–320
                0.46,
                0.48,  # рез 500 и остаток
                0.50,
                0.53,  # остаток ~500 и рез до 530
                0.70,
                0.72,  # рез 700 и остаток (740 по таблице не отдельный остаток)
                0.86,
                0.88,
                0.92,  # рез 900 и остатки
                1.00,
                1.02,
                1.08,
                1.20,  # рез 1020–1080 (1.02–1.08), целая 1.2
            ]
            # Берём максимальную стандартную ширину, не превышающую фактическую
            candidates = [w for w in STANDARD_WIDTHS if w <= width_m + 1e-6]
            if not candidates:
                # Слишком узкая или совсем нестандартная плита — игнорируем
                return
            snapped_width = max(candidates)

            # Рекурсивный вызов с притянутой шириной, чтобы сработали диапазоны выше.
            add_items(snapped_width, length_m, qty, load_code, length_dm_raw=length_dm_raw, line_idx=line_idx)
            return

        # Добавляем плиты в список и сохраняем точную ширину (длина с точностью 3 знака)
        length_rounded = round(float(length_m), 3)
        for _ in range(max(0, qty)):
            target.append(length_rounded)

            # Сохраняем точную ширину для этой плиты
            if target_name:
                key = (length_rounded, target_name)
                rt.plate_exact_widths[key] = round(width_m, 3)

        # Сохраняем нагрузку (если указана) в PLATE_LOAD_DETAILS и raw в PLATE_LENGTH_DM_RAW
        width_rounded = round(width_m, 3)
        ldr_norm = (length_dm_raw or "").strip()
        if load_code is not None and load_code > 0:
            key_new = (length_rounded, width_rounded, load_code, ldr_norm)
            rt.plate_load_details[key_new] = rt.plate_load_details.get(key_new, 0) + qty
            rt.plate_length_dm_raw[key_new] = ldr_norm
            if line_idx is not None and 0 <= line_idx < len(line_plate_load_details):
                _ld = line_plate_load_details[line_idx]
                _ld[key_new] = _ld.get(key_new, 0) + qty
            if line_idx is not None:
                _record_contribution(line_idx, length_rounded, width_rounded, float(load_code), ldr_norm)
        elif line_idx is not None:
            _record_contribution(line_idx, length_rounded, width_rounded, None, ldr_norm)

    # Список нераспознанных строк для отчёта пользователю
    unparsed_lines = []

    for line_idx, raw in enumerate(lines):
        parsed = False
        parsed_line = parse_line(raw)
        diag: dict[str, Any] = {
            "raw_input": raw,
            "parse_stage": parsed_line.stage,
            "recognized_by": "parser",
        }

        if not parsed_line.parsed:
            diag["validation_status"] = "failed"
            diag["reason_code"] = parsed_line.reason_code or "pattern_not_matched"
            diag["rejection_reason"] = parsed_line.reason_text or "строка не распознана"
            rt.last_parse_diagnostics.append(diag)
            unparsed_lines.append(f"{raw} (пропущено: {diag['rejection_reason']})")
            continue

        validation = validate_plate_values(parsed_line.width_m, parsed_line.length_m, parsed_line.qty)
        if not validation.ok:
            diag["validation_status"] = "failed"
            diag["reason_code"] = validation.reason_code
            diag["rejection_reason"] = validation.reason_text
            rt.last_parse_diagnostics.append(diag)
            unparsed_lines.append(f"{raw} (пропущено: {validation.reason_text})")
            continue

        add_items(
            parsed_line.width_m,
            parsed_line.length_m,
            parsed_line.qty,
            parsed_line.load_code,
            length_dm_raw=parsed_line.length_dm_raw,
            line_idx=line_idx,
        )
        parsed = True
        diag["validation_status"] = "ok"
        diag["normalized_input"] = raw
        rt.last_parse_diagnostics.append(diag)
        if parsed:
            continue

    _recompute_totals_from_lists()

    # Заполняем кэш номенклатуры один раз по всем ключам PLATE_LOAD_DETAILS + PLATE_LENGTH_DM_RAW
    _try_fill_plate_nomenclature_cache(fill_nomenclature_cache)

    # Логируем нераспознанные строки для отладки
    if unparsed_lines:
        logger.warning(
            f"Парсинг завершён с {len(unparsed_lines)} нераспознанными строками. "
            f"Всего строк обработано: {len(lines)}"
        )
        for i, line in enumerate(unparsed_lines[:5], 1):  # Показываем первые 5
            logger.debug(f"  Нераспознанная строка {i}: {line}")
        if len(unparsed_lines) > 5:
            logger.debug(f"  ... и ещё {len(unparsed_lines) - 5} строк")
    else:
        logger.info(f"Парсинг завершён успешно. Обработано строк: {len(lines)}")

    return unparsed_lines, line_contributions, line_plate_load_details
