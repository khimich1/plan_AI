"""Boundary adapters between app ``PlateOrder`` and canonical core ``PlateOrder``."""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING, Any

from core.domain.plate_order import PlateOrder as CorePlateOrder
from core.domain.plate_order import coerce_core_plate_order

if TYPE_CHECKING:
    from app.domain.models.plate_order import PlateOrder as AppPlateOrder


def to_core_order(order: Any) -> CorePlateOrder:
    """Strip app-only fields and return canonical core order."""
    if type(order) is CorePlateOrder:
        return order
    if hasattr(order, "to_dict"):
        data = dict(order.to_dict())
        data.pop("nomenclature_cache", None)
        return CorePlateOrder.from_dict(data)
    return coerce_core_plate_order(order)


def _copy_core_field_value(value: Any) -> Any:
    if isinstance(value, list):
        return list(value)
    if isinstance(value, dict):
        return dict(value)
    return value


def from_core_order(
    core: CorePlateOrder,
    *,
    nomenclature_cache: dict[tuple[float, float, int | float, str], dict[str, Any]] | None = None,
) -> AppPlateOrder:
    """Build app order from core; optionally override ``nomenclature_cache``."""
    from app.domain.models.plate_order import PlateOrder as AppPlateOrder

    if isinstance(core, AppPlateOrder) and nomenclature_cache is None:
        return core

    fields = {
        field.name: _copy_core_field_value(getattr(core, field.name))
        for field in dataclasses.fields(CorePlateOrder)
    }
    if nomenclature_cache is not None:
        cache = dict(nomenclature_cache)
    elif isinstance(core, AppPlateOrder):
        cache = dict(core.nomenclature_cache)
    else:
        cache = {}

    return AppPlateOrder(**fields, nomenclature_cache=cache)
