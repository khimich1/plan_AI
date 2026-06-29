from __future__ import annotations

from typing import Any


def order_data_from_kp_info(kp_info: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Собирает order_data для генераторов КП из kp_info['plates'].

    unit_price берётся из колонки или восстанавливается из discounted_price и скидки.
    Каноническая реализация для archive_service и повторного использования в боте.
    """
    plates = kp_info.get("plates") or []
    discount = kp_info.get("discount_percent") or 0
    factor = 1.0 - (discount / 100.0)
    if factor <= 0:
        factor = 1.0
    result: list[dict[str, Any]] = []
    for plate in plates:
        unit_price = plate.get("unit_price")
        discounted_price = plate.get("discounted_price")
        resolved_unit_price: float | None = None
        if isinstance(unit_price, (int, float)) and unit_price > 0:
            resolved_unit_price = float(unit_price)
        elif isinstance(discounted_price, (int, float)) and discounted_price > 0:
            resolved_unit_price = float(discounted_price) / factor

        qty = plate.get("qty") or 0
        total_weight = plate.get("total_weight")
        unit_weight = plate.get("unit_weight")
        weight = (
            total_weight
            if total_weight is not None and total_weight > 0
            else (unit_weight or 0) * qty
        )
        item: dict[str, Any] = {
            "name": plate.get("plate_name") or "",
            "length_m": plate.get("length_m") or 0,
            "width_m": plate.get("width_m") or 0,
            "qty": qty,
            "load_class": plate.get("load_class") or 800,
            "weight": weight or 0,
        }
        if resolved_unit_price is not None:
            item["unit_price"] = resolved_unit_price
        result.append(item)
    return result
