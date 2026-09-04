# -*- coding: utf-8 -*-
"""Physical dimensions, cut pricing, and plate dimension parsing helpers."""

import logging

logger = logging.getLogger(__name__)

TRACK_LENGTH_M = 101.0
TRACK_WIDTH_M = 1.2

LONG_CUT_PRICE_PER_M = 460.0  # Продольный рез, руб/пог.м
TRANSVERSE_CUT_PRICE = 1200.0  # Поперечный (или скошенный) рез, руб/шт
MIN_BILLABLE_TRIM_MM = 20  # Остаток/отход ≤ 20 мм не тарифицируется и не требует реза
PLATE_WIDTH_MATCH_TOLERANCE_MM = 10  # 10 мм = plate mark tolerance (1 cm); 20 мм MIN_BILLABLE = uncut kerf

WEIGHT_KG_PER_DM2 = 2.83333333


def length_dm_to_m(Ldm_str: str) -> float:
    """
    Переводит строку длины из марки ПБ в метры.
    - Нет запятой/точки (например "38") → номинал в дм: длина = номинал / 10 м (38 → 3.8 м).
    - Есть запятая/точка ("38,0", "75,5") → точное значение: парсим в целые мм, затем одно деление на 1000.
    """
    s = (Ldm_str or "").strip().replace(" ", "")
    if "," in s or "." in s:
        # Ветка с запятой/точкой: разбираем строку без float, считаем длину в целых мм
        parts = s.replace(",", ".").split(".")
        if len(parts) != 2:
            try:
                return round(float(s.replace(",", ".")) / 10.0, 3)
            except ValueError:
                return 0.0
        try:
            int_part = int(parts[0].strip())
            frac_str = (parts[1].strip() or "0")[:3]  # не более 3 знаков
            frac_part = int(frac_str) if frac_str else 0
            denom = 10 ** len(frac_str)  # 1, 10, 100
            # Длина в мм: value_dm = int_part + frac_part/denom, length_mm = value_dm * 100
            # Целочисленно: (int_part * denom + frac_part) * 100 // denom
            length_mm = (int_part * denom + frac_part) * 100 // denom
        except ValueError:
            return 0.0
        if length_mm < 0:
            return 0.0
        return round(length_mm / 1000.0, 3)
    try:
        nominal_dm = int(float(s))
    except ValueError:
        return 0.0
    # Номинал в дм → длина в метрах без вычета 20 мм (69 → 6.9 м)
    return round(nominal_dm / 10.0, 3)


def normalize_dimension(value_str: str) -> float:
    """
    Нормализует размер из строки в дециметры.

    Примеры:
    - "5,30" -> 5.3
    - "530" -> 5.3 (введены мм без разделителя)
    - "6.65" -> 6.65
    """
    clean = value_str.strip().replace(" ", "").replace(",", ".")
    try:
        val = float(clean)
    except ValueError:
        return 0.0

    # Диапазон значений, которые чаще всего вводят как мм без запятой:
    # 530 -> 5.30 дм, 665 -> 6.65 дм.
    if 20.0 < val < 1000.0:
        val = val / 100.0
        logger.debug(
            "normalize_dimension: применили /100 для значения, похожего на мм: %s -> %s",
            value_str.strip(),
            val,
        )
    return val


def parse_pb_width_to_m(width_str: str) -> float:
    """
    Парсит ширину из марки ПБ/ПК (часть W в L-W-N) в метры.

    Правила:
    - 0.2/0.3 (или 0,2/0,3) считаются уже метрами (спец-ленты).
    - Остальные значения интерпретируются как дециметры и делятся на 10.
    """
    width_raw = normalize_dimension(width_str)
    if width_raw <= 0:
        return 0.0
    if abs(width_raw - 0.3) < 1e-6 or abs(width_raw - 0.2) < 1e-6:
        return round(width_raw, 3)
    return round(width_raw / 10.0, 3)
