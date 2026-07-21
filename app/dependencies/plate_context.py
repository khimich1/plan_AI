from __future__ import annotations

from fastapi import HTTPException, Request, status

from core.plate_order_context import PlateOrderContext


def get_plate_order_context(request: Request) -> PlateOrderContext:
    """Контекст заказа, установленный ``PlateMutableRuntimeIsolationMiddleware``."""
    ctx = getattr(request.state, "plate_order_ctx", None)
    if not isinstance(ctx, PlateOrderContext):
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Plate order context not initialized",
        )
    return ctx
