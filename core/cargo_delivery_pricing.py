#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт услуги по доставке грузов в коммерческих предложениях.

Поле logistics_cost / «стоимость рейса» трактуется как цена одного рейса.
Итог: стоимость_рейса × ceil(масса_груза_кг / грузоподъёмность).
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, List, Mapping, Optional, Sequence

try:
    from .kp_plate_weight import resolve_kp_line_weight_kg
except ImportError:
    from kp_plate_weight import resolve_kp_line_weight_kg

# Максимальная масса груза на один рейс (кг), согласно ТЗ.
CARGO_DELIVERY_TRUCK_CAPACITY_KG: float = 18600.0
FBS_LM_TRUCK_CAPACITY_KG: float = 20_000.0
FBS_LM_PRODUCT_TYPES = frozenset({"fbs", "steps", "marches"})
FBS_LM_PRODUCT_KINDS = frozenset({"fbs", "step", "march"})
_KIND_TO_WEIGHT_TYPE = {"fbs": "fbs", "step": "steps", "march": "marches"}


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


@dataclass(frozen=True)
class WeightedDeliveryBreakdown:
    cargo_kg: float
    trips: int  # 0 если pending_marks не пуст
    pending_marks: tuple[str, ...]

    @property
    def ready(self) -> bool:
        return not self.pending_marks


def _fbs_lm_trips_count(cargo_weight_kg: float) -> int:
    w = max(0.0, float(cargo_weight_kg or 0.0))
    if w <= 0:
        return 0
    return int(math.ceil(w / FBS_LM_TRUCK_CAPACITY_KG))


def _line_qty(item: Mapping[str, Any]) -> int:
    try:
        qty = int(item.get("qty") or 0)
    except (TypeError, ValueError):
        return 0
    return max(0, qty)


def fbs_lm_catalog_type(item: Mapping[str, Any]) -> Optional[str]:
    """Тип для справочника весов: product_type либо product_kind (step→steps)."""
    product_type = str(item.get("product_type") or "").strip().lower()
    if product_type in FBS_LM_PRODUCT_TYPES:
        return product_type
    kind = str(item.get("product_kind") or "").strip().lower()
    if kind in FBS_LM_PRODUCT_KINDS:
        return _KIND_TO_WEIGHT_TYPE[kind]
    return None


def compute_fbs_lm_delivery(
    order_data: Sequence[Mapping[str, Any]],
    *,
    trip_cost: float,
    catalog_db_path: str,
) -> WeightedDeliveryBreakdown:
    """Σ масса строк fbs/steps/marches → ceil(кг/20 000).

    Тип строки: product_type ∈ FBS_LM_PRODUCT_TYPES (mixed-builder)
    или product_kind ∈ FBS_LM_PRODUCT_KINDS (mono/legacy).
    Марка без веса → pending_marks, котёл не готов (как у свай).
    """
    del trip_cost  # сумма считается в calculate_total_cost, как у свай
    from core.product_weight_catalog import resolve_product_weight_kg

    cargo_kg = 0.0
    pending_display: list[str] = []
    seen_pending: set[str] = set()
    for item in order_data:
        catalog_type = fbs_lm_catalog_type(item)
        if catalog_type is None:
            continue
        mark = str(item.get("mark") or item.get("name") or "").strip()
        qty = _line_qty(item)
        if not mark or qty <= 0:
            continue
        unit_kg = resolve_product_weight_kg(mark, catalog_type, catalog_db_path)
        if unit_kg is None:
            if mark not in seen_pending:
                seen_pending.add(mark)
                pending_display.append(mark)
            continue
        cargo_kg += float(unit_kg) * qty

    pending_marks = tuple(pending_display)
    trips = 0 if pending_marks else _fbs_lm_trips_count(cargo_kg)
    return WeightedDeliveryBreakdown(
        cargo_kg=cargo_kg,
        trips=trips,
        pending_marks=pending_marks,
    )
