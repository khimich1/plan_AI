#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Единый расчёт массы строки КП: plate_weights по длине/ширине, затем approximate_weight_kg.

Нагрузка в марке на массу не влияет (в БД опорная запись ПБ {L}-12-8).
Поле order_data['weight'] — суммарный вес строки (на все шт.), см. save_kp_to_db.
"""

from __future__ import annotations

import logging
from typing import Any, Mapping

from .config_and_data import approximate_weight_kg
from .plate_weights_db import get_plate_weight_kg_by_dimensions

logger = logging.getLogger(__name__)


def resolve_kp_line_weight_kg(item: Mapping[str, Any]) -> tuple[float, float]:
    """
    Возвращает (unit_weight_kg, total_weight_kg) для позиции заказа.

    Приоритет: get_plate_weight_kg_by_dimensions → approximate_weight_kg.
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

    unit: float | None = None
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
                "approximate_weight_kg failed for length_m=%s width_m=%s",
                length_m,
                width_m,
            )
            unit = 0.0

    total = unit * qty
    return (unit, total)
