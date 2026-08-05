from __future__ import annotations

from typing import Any


def _restore_unit_price(plate: dict[str, Any], discount: float) -> float:
    unit_price = plate.get("unit_price")
    if unit_price is not None and isinstance(unit_price, (int, float)) and unit_price > 0:
        return float(unit_price)
    factor = 1.0 - (discount / 100.0)
    if factor <= 0:
        factor = 1.0
    discounted_price = plate.get("discounted_price") or 0
    return float(discounted_price) / factor


def order_data_from_kp_piles(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """Build pile order_data from kp_info['piles']."""
    piles = kp_info.get("piles") or []
    discount = float(kp_info.get("discount_percent") or 0)
    result: list[dict[str, Any]] = []
    for pile in piles:
        unit_price = _restore_unit_price(pile, discount)
        mark = str(pile.get("mark") or "").strip()
        result.append(
            {
                "product_kind": "pile",
                "name": mark,
                "mark": mark,
                "concrete_grade": str(pile.get("concrete_grade") or "B25").strip(),
                "qty": int(pile.get("qty") or 0),
                "unit_price": unit_price,
            }
        )
    return result


def order_data_from_kp_marches(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """Build march order_data from kp_info['marches']."""
    marches = kp_info.get("marches") or []
    discount = float(kp_info.get("discount_percent") or 0)
    result: list[dict[str, Any]] = []
    for march in marches:
        unit_price = _restore_unit_price(march, discount)
        mark = str(march.get("mark") or "").strip()
        result.append(
            {
                "product_kind": "march",
                "name": mark,
                "mark": mark,
                "concrete_grade": str(march.get("concrete_grade") or "B25").strip(),
                "qty": int(march.get("qty") or 0),
                "unit_price": unit_price,
            }
        )
    return result


def order_data_from_kp_bridge_piles(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """Build bridge-pile order_data from kp_info['bridge_piles']."""
    items = kp_info.get("bridge_piles") or []
    discount = float(kp_info.get("discount_percent") or 0)
    result: list[dict[str, Any]] = []
    for item in items:
        unit_price = _restore_unit_price(item, discount)
        mark = str(item.get("mark") or "").strip()
        result.append(
            {
                "product_kind": "bridge_pile",
                "name": mark,
                "mark": mark,
                "concrete_grade": str(item.get("concrete_grade") or "B25").strip(),
                "qty": int(item.get("qty") or 0),
                "unit_price": unit_price,
            }
        )
    return result



def order_data_from_kp_fbs(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """Build FBS order_data from kp_info['fbs']."""
    items = kp_info.get("fbs") or []
    discount = float(kp_info.get("discount_percent") or 0)
    result: list[dict[str, Any]] = []
    for item in items:
        unit_price = _restore_unit_price(item, discount)
        mark = str(item.get("mark") or "").strip()
        result.append(
            {
                "product_kind": "fbs",
                "name": mark,
                "mark": mark,
                "concrete_grade": str(item.get("concrete_grade") or "B25").strip(),
                "qty": int(item.get("qty") or 0),
                "unit_price": unit_price,
            }
        )
    return result


def order_data_from_kp_steps(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """Build step order_data from kp_info['steps']."""
    steps = kp_info.get("steps") or []
    discount = float(kp_info.get("discount_percent") or 0)
    result: list[dict[str, Any]] = []
    for step in steps:
        unit_price = _restore_unit_price(step, discount)
        mark = str(step.get("mark") or "").strip()
        result.append(
            {
                "product_kind": "step",
                "name": mark,
                "mark": mark,
                "qty": int(step.get("qty") or 0),
                "unit_price": unit_price,
            }
        )
    return result


def order_data_from_kp_plates(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """Build plate order_data from kp_info['plates']."""
    plates = kp_info.get("plates") or []
    discount = float(kp_info.get("discount_percent") or 0)
    result: list[dict[str, Any]] = []
    for plate in plates:
        unit_price = _restore_unit_price(plate, discount)
        qty = plate.get("qty") or 0
        total_weight = plate.get("total_weight")
        unit_weight = plate.get("unit_weight")
        weight = (
            total_weight
            if total_weight is not None and total_weight > 0
            else (unit_weight or 0) * qty
        )
        result.append(
            {
                "name": plate.get("plate_name") or "",
                "length_m": plate.get("length_m") or 0,
                "width_m": plate.get("width_m") or 0,
                "qty": qty,
                "load_class": plate.get("load_class") or 800,
                "unit_price": unit_price,
                "weight": weight or 0,
            }
        )
    return result


def order_data_from_kp_info(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Собирает order_data для генераторов КП из kp_info.

    unit_price берётся из колонки или восстанавливается из discounted_price и скидки.
    Каноническая реализация для archive_service и повторного использования в боте.
    """
    product_type = str(kp_info.get("product_type") or "plates").lower()
    if product_type == "piles" or kp_info.get("piles"):
        return order_data_from_kp_piles(kp_info)
    if product_type == "bridge_piles" or kp_info.get("bridge_piles"):
        return order_data_from_kp_bridge_piles(kp_info)
    if product_type == "fbs" or kp_info.get("fbs"):
        return order_data_from_kp_fbs(kp_info)
    if product_type == "marches" or kp_info.get("marches"):
        return order_data_from_kp_marches(kp_info)
    if product_type == "steps" or kp_info.get("steps"):
        return order_data_from_kp_steps(kp_info)
    return order_data_from_kp_plates(kp_info)
