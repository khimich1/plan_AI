from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from core.domain.plate_order import PlateOrder as CorePlateOrder
from core.domain.plate_order import parse_load_key_from_list

from app.domain.adapters.plate_order import from_core_order


LoadKey = tuple[float, float, int | float, str]
ExactWidthKey = tuple[float, str]


@dataclass
class PlateOrder(CorePlateOrder):
    """App-layer order: canonical core fields + ``nomenclature_cache`` for commercial flows."""

    nomenclature_cache: dict[LoadKey, dict[str, Any]] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        data = super().to_dict()
        data["nomenclature_cache"] = [
            [list(k), v] for k, v in self.nomenclature_cache.items()
        ]
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PlateOrder:
        cache_entries = data.get("nomenclature_cache", [])
        core_data = {key: value for key, value in data.items() if key != "nomenclature_cache"}
        core = CorePlateOrder.from_dict(core_data)
        nomenclature_cache = {
            parse_load_key_from_list(k): dict(v) for k, v in cache_entries
        }
        return from_core_order(core, nomenclature_cache=nomenclature_cache)

    @classmethod
    def from_legacy(cls, legacy_order: Any) -> PlateOrder:
        if isinstance(legacy_order, cls):
            return legacy_order
        if hasattr(legacy_order, "to_dict"):
            return cls.from_dict(legacy_order.to_dict())
        raise TypeError("Unsupported legacy plate order")

    @classmethod
    def from_orders_2d(cls, orders_2d: list[dict[str, Any]]) -> PlateOrder:
        return from_core_order(CorePlateOrder.from_orders_2d(orders_2d))

    def recompute_totals(self) -> None:
        super().recompute_totals()
