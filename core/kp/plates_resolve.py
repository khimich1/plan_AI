"""Сбор позиций КП для архива и генерации PDF/XLSX."""

from __future__ import annotations

from typing import Any

from core.kp_db_plates_queries import get_completed_plates_for_kp


def aggregate_completed_plates(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Схлопывает строки completed_plates в позиции как в kp_plates (сумма qty)."""
    buckets: dict[tuple[Any, ...], dict[str, Any]] = {}
    order_keys: list[tuple[Any, ...]] = []

    for row in rows:
        qty = int(row.get("qty") or 0)
        if qty <= 0:
            continue
        key = (
            row.get("plate_name") or "",
            float(row.get("length_m") or 0),
            float(row.get("width_m") or 0),
            int(row.get("load_class") or 800),
        )
        if key not in buckets:
            buckets[key] = {
                "plate_name": key[0],
                "length_m": key[1],
                "width_m": key[2],
                "load_class": key[3],
                "qty": 0,
                "status": "выполнено",
            }
            order_keys.append(key)
        buckets[key]["qty"] += qty

    result: list[dict[str, Any]] = []
    for index, key in enumerate(order_keys, start=1):
        plate = dict(buckets[key])
        plate["position_number"] = index
        result.append(plate)
    return result


def resolve_plates_for_kp_documents(
    active_plates: list[dict[str, Any]],
    *,
    kp_id: int,
    db_path: str,
) -> list[dict[str, Any]]:
    """
    Позиции для документов: сначала kp_plates с qty > 0,
    иначе агрегированные completed_plates (полностью выполненное КП).
    """
    active = [plate for plate in active_plates if int(plate.get("qty") or 0) > 0]
    if active:
        return active
    completed = get_completed_plates_for_kp(kp_id, db_path)
    return aggregate_completed_plates(completed)
