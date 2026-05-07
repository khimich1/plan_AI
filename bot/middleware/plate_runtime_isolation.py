"""Изоляция мутабельного заказа плит на время обработки апдейта (S1)."""

from __future__ import annotations

from typing import Any, Awaitable, Callable, Dict

from aiogram import BaseMiddleware
from aiogram.types import TelegramObject

from core.plate_runtime_state import fresh_plate_mutable_request_scope


class PlateMutableRuntimeIsolationMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any],
    ) -> Any:
        with fresh_plate_mutable_request_scope():
            return await handler(event, data)
