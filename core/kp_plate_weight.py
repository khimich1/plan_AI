#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Единый расчёт массы строки КП с переключаемым источником веса.

По умолчанию используется формула из config_and_data (режим WEIGHT_SOURCE="formula").
Legacy-путь через plate_weights доступен при WEIGHT_SOURCE="plate_weights".

Нагрузка в марке на массу не влияет (в legacy-режиме БД хранит опорную запись ПБ {L}-12-8).
Поле order_data['weight'] — суммарный вес строки (на все шт.), см. save_kp_to_db.
"""

from __future__ import annotations

import logging
from typing import Any, Mapping

from .config_and_data import WEIGHT_SOURCE, approximate_weight_kg
from .plate_weights_db import get_plate_weight_kg_by_dimensions

logger = logging.getLogger(__name__)


def resolve_kp_line_weight_kg(item: Mapping[str, Any]) -> tuple[float, float]:
    """
    Возвращает (unit_weight_kg, total_weight_kg) для позиции заказа.

    Режимы:
    - formula (default): approximate_weight_kg
    - plate_weights (legacy): get_plate_weight_kg_by_dimensions -> fallback approximate_weight_kg
    """
    try:
        qty = float(item.get("qty") or 0)
    except (TypeError, ValueError):
        qty = 0.0
    try:
        length_m = float(item.get("length_m") or 0)
    except (TypeError, ValueError):
        length_m = 0.0
    try:
        width_m = float(item.get("width_m") or 0)
    except (TypeError, ValueError):
        width_m = 0.0

    source = WEIGHT_SOURCE if WEIGHT_SOURCE in {"formula", "plate_weights"} else "formula"
    unit: float | None = None

    if source == "plate_weights":
        try:
            unit = get_plate_weight_kg_by_dimensions(length_m, width_m)
        except Exception:
            logger.exception(
                "get_plate_weight_kg_by_dimensions failed for length_m=%s width_m=%s",
                length_m,
                width_m,
            )
            unit = None

    if unit is None:
        try:
            unit = approximate_weight_kg(length_m, width_m)
        except Exception:
            logger.exception(
                "approximate_weight_kg failed for source=%s length_m=%s width_m=%s",
                source,
                length_m,
                width_m,
            )
            unit = 0.0

    total = unit * qty
    return (unit, total)
