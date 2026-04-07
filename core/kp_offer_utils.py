from __future__ import annotations

from typing import Any, Callable, Mapping


TRANSPORT_LINE_NAME = "Транспортные расходы"
TRANSPORT_LINE_TYPE = "transport"
TRANSPORT_UNIT = "час"


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def format_offer_quantity(value: Any) -> str:
    qty = _safe_float(value)
    if qty.is_integer():
        return str(int(qty))
    return f"{qty:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def get_offer_item_unit(item: Mapping[str, Any]) -> str:
    unit = str(item.get("unit") or "").strip()
    return unit or "шт"


def is_transport_offer_item(item: Mapping[str, Any]) -> bool:
    line_type = str(item.get("line_type") or "").strip().lower()
    if line_type == TRANSPORT_LINE_TYPE:
        return True

    name = str(item.get("name") or "").strip().lower()
    unit = get_offer_item_unit(item).lower()
    return name == TRANSPORT_LINE_NAME.lower() and unit == TRANSPORT_UNIT


def is_weightless_offer_item(item: Mapping[str, Any]) -> bool:
    return bool(item.get("exclude_from_weight")) or is_transport_offer_item(item)


def build_transport_order_item(
    transport_hours: Any,
    transport_price_per_hour: Any,
) -> dict[str, Any] | None:
    hours = _safe_float(transport_hours)
    price_per_hour = _safe_float(transport_price_per_hour)
    if hours <= 0 or price_per_hour <= 0:
        return None

    return {
        "name": TRANSPORT_LINE_NAME,
        "length_m": 0,
        "width_m": 0,
        "qty": hours,
        "unit": TRANSPORT_UNIT,
        "load_class": 800,
        "unit_price": price_per_hour,
        "weight": 0,
        "line_type": TRANSPORT_LINE_TYPE,
        "exclude_from_discount": True,
        "exclude_from_weight": True,
    }


def extract_transport_from_order_data(order_data: list[Mapping[str, Any]]) -> tuple[float | None, float | None]:
    for item in order_data:
        if not is_transport_offer_item(item):
            continue
        hours = _safe_float(item.get("qty"))
        price_per_hour = _safe_float(item.get("unit_price"))
        if hours > 0 and price_per_hour > 0:
            return hours, price_per_hour
    return None, None


def append_transport_to_order_data(
    order_data: list[Mapping[str, Any]],
    transport_hours: Any,
    transport_price_per_hour: Any,
) -> list[dict[str, Any]]:
    result = [dict(item) for item in order_data if not is_transport_offer_item(item)]
    transport_item = build_transport_order_item(transport_hours, transport_price_per_hour)
    if transport_item:
        result.append(transport_item)
    return result


def _resolve_base_unit_price(
    item: Mapping[str, Any],
    fallback_price_getter: Callable[[float, float, int], float] | None = None,
) -> float:
    unit_price = item.get("unit_price")
    if unit_price is not None:
        return _safe_float(unit_price)

    if fallback_price_getter is None:
        return 0.0

    length_m = _safe_float(item.get("length_m"))
    width_m = _safe_float(item.get("width_m"))
    load_class = int(_safe_float(item.get("load_class"), default=800))
    return _safe_float(fallback_price_getter(length_m, width_m, load_class))


def get_offer_item_price_with_discount(
    item: Mapping[str, Any],
    discount_percent: float = 0,
    fallback_price_getter: Callable[[float, float, int], float] | None = None,
) -> float:
    unit_price = _resolve_base_unit_price(item, fallback_price_getter)
    if item.get("exclude_from_discount"):
        return unit_price
    return unit_price * (1 - discount_percent / 100)


def calculate_offer_totals(
    order_data: list[Mapping[str, Any]],
    discount_percent: float = 0,
    fallback_price_getter: Callable[[float, float, int], float] | None = None,
) -> dict[str, float]:
    total_qty = 0.0
    total_cost_with_vat = 0.0

    for item in order_data:
        qty = _safe_float(item.get("qty"))
        unit_price = get_offer_item_price_with_discount(
            item,
            discount_percent=discount_percent,
            fallback_price_getter=fallback_price_getter,
        )
        total_qty += qty
        total_cost_with_vat += unit_price * qty

    subtotal = round(total_cost_with_vat / 1.22, 2)
    vat_amount = round(total_cost_with_vat - subtotal, 2)
    total_with_vat = round(total_cost_with_vat, 2)

    return {
        "total_qty": total_qty,
        "subtotal": subtotal,
        "vat_amount": vat_amount,
        "total_with_vat": total_with_vat,
    }
