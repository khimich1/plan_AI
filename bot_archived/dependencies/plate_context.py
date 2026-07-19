from __future__ import annotations

from typing import Any

from aiogram.filters import Filter

from core.plate_order_context import PlateOrderContext


def get_plate_order_context(data: dict[str, Any]) -> PlateOrderContext:
    """Контекст заказа из aiogram ``data`` (``PlateMutableRuntimeIsolationMiddleware``)."""
    ctx = data.get("plate_order_ctx")
    if not isinstance(ctx, PlateOrderContext):
        raise RuntimeError("Plate order context not initialized")
    return ctx


class PlateOrderContextDep(Filter):
    """Aiogram filter: inject ``plate_order_ctx`` via ``get_plate_order_context``."""

    async def __call__(self, **kwargs: Any) -> dict[str, PlateOrderContext]:
        return {"plate_order_ctx": get_plate_order_context(kwargs)}
