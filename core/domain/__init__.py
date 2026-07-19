"""Доменные сущности core (legacy-совместимые агрегаты)."""

from .plate_order import (
    PlateOrder,
    get_current_plate_order,
    normalize_load_code,
)

__all__ = [
    "PlateOrder",
    "get_current_plate_order",
    "normalize_load_code",
]
