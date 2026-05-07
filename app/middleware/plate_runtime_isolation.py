"""Изоляция мутабельного заказа плит на время HTTP-запроса (S1)."""

from __future__ import annotations

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import Response

from core.plate_runtime_state import fresh_plate_mutable_request_scope


class PlateMutableRuntimeIsolationMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next) -> Response:
        with fresh_plate_mutable_request_scope():
            return await call_next(request)
