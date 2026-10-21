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
from typing import Any, Callable, List, Mapping, Optional, Sequence

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


def document_cargo_weight_kg(
    order_data: Sequence[Mapping[str, Any]],
    *,
    pile_catalog_db_path: str | None = None,
    weight_catalog_db_path: str | None = None,
) -> float:
    """Физическая масса заказа для подвала КП: штуки × вес строки.

    Плиты — ``resolve_kp_line_weight_kg``. Сваи и мостовые — ``pile_catalog``
    (суффикс нагрузки отрезается). ФБС, марши и ступени — ``product_weight_catalog``.
    Составные сваи и марки без веса в сумму не входят.

    Это не остаток рейса: полная машина входит целиком.
    """
    from core.commercial_offer_layout import line_product_type

    pile_path = pile_catalog_db_path or _default_plita_db_path()
    weight_path = weight_catalog_db_path or pile_path
    pile_lookup = _pile_weight_lookup(pile_path)
    total = 0.0
    for item in order_data:
        kind = line_product_type(item)
        if kind == "plates":
            _, line_kg = resolve_kp_line_weight_kg(item)
            total += line_kg
            continue
        qty = _line_qty(item)
        if qty <= 0:
            continue
        if kind in {"piles", "bridge_piles"}:
            unit_kg = _pile_unit_weight_kg(item, pile_lookup)
        elif kind in {"fbs", "steps", "marches"}:
            unit_kg = _fbs_lm_unit_weight_kg(item, kind, weight_path)
        else:
            unit_kg = None
        if unit_kg is None or unit_kg <= 0:
            continue
        total += float(unit_kg) * qty
    return total


def _default_plita_db_path() -> str:
    try:
        from core.config.settings import get_settings

        return str(get_settings().plita_db_path)
    except Exception:
        return ""


def _pile_weight_lookup(db_path: str):
    if not db_path:
        return None
    from core.pile_catalog import load_pile_catalog, resolve_catalog_for_mark

    entries = load_pile_catalog(db_path)
    return lambda mark: resolve_catalog_for_mark(mark, entries)


def _pile_unit_weight_kg(item: Mapping[str, Any], lookup) -> float | None:
    if lookup is None:
        return None
    mark = str(item.get("mark") or item.get("name") or "").strip()
    if not mark:
        return None
    entry = lookup(mark)
    if entry is None:
        return None
    return float(entry.weight_kg)


def _fbs_lm_unit_weight_kg(item: Mapping[str, Any], kind: str, db_path: str) -> float | None:
    if not db_path:
        return None
    from core.product_weight_catalog import resolve_product_weight_kg

    mark = str(item.get("mark") or item.get("name") or "").strip()
    if not mark:
        return None
    return resolve_product_weight_kg(mark, kind, db_path)


def total_order_cargo_weight_kg(
    order_data: List[Mapping[str, Any]],
    product_types: set[str] | None = None,
) -> float:
    """Суммарная масса по resolve_kp_line_weight_kg (габариты плиты).

    product_types=None — все позиции (обратная совместимость).
    Иначе учитываются только строки с product_type из множества;
    отсутствие product_type трактуется как «plates» (legacy mono).

    Строка «общим весом» в PDF/XLSX — kp_document_total_weight_kg:
    плиты считаются здесь же, сваи и ФБС/ЛС/ЛМ — из своих справочников.
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


def _pile_catalog_lookup(db_path: str | None) -> Callable[[str], Any]:
    """Тот же resolve_catalog_for_mark, что у расчёта рейсов свай."""
    from core.commercial_pricing import _resolve_pile_catalog_db_path
    from core.pile_catalog import load_pile_catalog, resolve_catalog_for_mark

    catalog_path = _resolve_pile_catalog_db_path(db_path, "")
    entries = load_pile_catalog(catalog_path)

    def lookup(mark: str) -> Any:
        return resolve_catalog_for_mark(mark, entries)

    return lookup


def _resolved_weight_catalog_path(db_path: str | None) -> str:
    from core.commercial_pricing import _resolve_weight_catalog_db_path

    return _resolve_weight_catalog_db_path(db_path, "")


def _catalog_line_kg(item: Mapping[str, Any], unit_kg: float | None) -> float:
    qty = _line_qty(item)
    if unit_kg is None or qty <= 0:
        return 0.0
    unit = float(unit_kg)
    if unit <= 0:
        return 0.0
    return unit * qty


def kp_document_line_weight_kg(
    item: Mapping[str, Any],
    *,
    pile_lookup: Callable[[str], Any] | None = None,
    weight_catalog_db_path: str | None = None,
) -> float:
    """Масса одной строки для итога «общим весом» в PDF/XLSX.

    Плиты — resolve_kp_line_weight_kg. Сваи и мостовые — weight_kg справочника
    свай × qty. ФБС/ЛС/ЛМ — resolve_product_weight_kg × qty. Составные сваи
    и типы без готового расчёта — 0. Колонка «Вес(кг)» у плит по-прежнему
    считается через resolve_kp_line_weight_kg.
    """
    from core.commercial_offer_layout import line_product_type
    from core.pile_trip_pricing import PILE_TRIP_PRODUCT_TYPES

    product_type = line_product_type(item)
    if product_type == "plates":
        _, line_kg = resolve_kp_line_weight_kg(item)
        return float(line_kg)

    if product_type in PILE_TRIP_PRODUCT_TYPES:
        if pile_lookup is None:
            return 0.0
        mark = str(item.get("mark") or item.get("name") or "").strip()
        if not mark:
            return 0.0
        entry = pile_lookup(mark)
        unit_kg = None if entry is None else getattr(entry, "weight_kg", None)
        return _catalog_line_kg(item, unit_kg)

    catalog_type = fbs_lm_catalog_type(item)
    if catalog_type is None or not weight_catalog_db_path:
        return 0.0
    mark = str(item.get("mark") or item.get("name") or "").strip()
    if not mark:
        return 0.0
    from core.product_weight_catalog import resolve_product_weight_kg

    return _catalog_line_kg(
        item,
        resolve_product_weight_kg(mark, catalog_type, weight_catalog_db_path),
    )


def kp_document_total_weight_kg(
    order_data: Sequence[Mapping[str, Any]],
    *,
    pile_catalog_db_path: str | None = None,
    weight_catalog_db_path: str | None = None,
) -> float:
    """Сумма кг для PDF/XLSX. Справочник открывается только если в заказе есть его тип."""
    from core.commercial_offer_layout import line_product_type
    from core.pile_trip_pricing import PILE_TRIP_PRODUCT_TYPES

    product_types = [line_product_type(item) for item in order_data]
    pile_lookup = None
    if any(product_type in PILE_TRIP_PRODUCT_TYPES for product_type in product_types):
        pile_lookup = _pile_catalog_lookup(pile_catalog_db_path)

    resolved_weight_path = None
    if any(fbs_lm_catalog_type(item) is not None for item in order_data):
        resolved_weight_path = _resolved_weight_catalog_path(weight_catalog_db_path)

    total = 0.0
    for item in order_data:
        total += kp_document_line_weight_kg(
            item,
            pile_lookup=pile_lookup,
            weight_catalog_db_path=resolved_weight_path,
        )
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
