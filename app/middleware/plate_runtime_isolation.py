"""Изоляция мутабельного заказа плит на время HTTP-запроса (S1 / A1-001)."""

from __future__ import annotations

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import Response

from core.plate_order_context import PlateOrderContext


class PlateMutableRuntimeIsolationMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next) -> Response:
        ctx = PlateOrderContext.fresh_empty()
        request.state.plate_order_ctx = ctx
        with ctx.bound():
            return await call_next(request)
