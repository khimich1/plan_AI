#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт услуги по доставке грузов в коммерческих предложениях.

Поле logistics_cost / «стоимость рейса» трактуется как цена одного рейса.
Итог: стоимость_рейса × ceil(масса_груза_кг / грузоподъёмность).
"""

from __future__ import annotations

import math
from typing import Any, List, Mapping

try:
    from .kp_plate_weight import resolve_kp_line_weight_kg
except ImportError:
    from kp_plate_weight import resolve_kp_line_weight_kg

# Максимальная масса груза на один рейс (кг), согласно ТЗ.
CARGO_DELIVERY_TRUCK_CAPACITY_KG: float = 18600.0


def total_order_cargo_weight_kg(
    order_data: List[Mapping[str, Any]],
    product_types: set[str] | None = None,
) -> float:
    """Суммарная масса позиций заказа в кг (та же логика, что в PDF/XLSX КП).

    product_types=None — все позиции (обратная совместимость).
    Иначе учитываются только строки с product_type из множества;
    отсутствие product_type трактуется как «plates» (legacy mono).
    """
    total = 0.0
    for item in order_data:
        if product_types is not None:
            line_type = item.get("product_type") or "plates"
            if line_type not in product_types:
                continue
        _, line_kg = resolve_kp_line_weight_kg(item)
        total += line_kg
    return total


def cargo_delivery_trips_count(cargo_weight_kg: float) -> int:
    """Число рейсов: округление вверх массы к кратности грузоподъёмности."""
    w = max(0.0, float(cargo_weight_kg or 0.0))
    if w <= 0:
        return 0
    return int(math.ceil(w / CARGO_DELIVERY_TRUCK_CAPACITY_KG))


def delivery_service_charge_rub(trip_cost_rub: float, cargo_weight_kg: float) -> float:
    """Итоговая сумма строки «услуга по доставке грузов» (без НДС в базе плит)."""
    trip = max(0.0, float(trip_cost_rub or 0.0))
    n = cargo_delivery_trips_count(cargo_weight_kg)
    return round(trip * n, 2)
