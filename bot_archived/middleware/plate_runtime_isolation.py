"""Изоляция мутабельного заказа плит на время обработки апдейта (S1 / A1-001)."""

from __future__ import annotations

from typing import Any, Awaitable, Callable, Dict

from aiogram import BaseMiddleware
from aiogram.types import TelegramObject

from core.plate_order_context import PlateOrderContext


class PlateMutableRuntimeIsolationMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any],
    ) -> Any:
        ctx = PlateOrderContext.fresh_empty()
        data["plate_order_ctx"] = ctx
        with ctx.bound():
            return await handler(event, data)
