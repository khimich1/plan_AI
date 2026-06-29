"""Per-router role authorization middleware."""

from __future__ import annotations

from typing import Any, Awaitable, Callable

from aiogram import BaseMiddleware
from aiogram.types import CallbackQuery, Message, TelegramObject

from bot.security.audit import log_bot_security_event
from bot.security.users import BotUser, has_role

_FORBIDDEN_TEXT = "⛔ У вас нет прав для этого действия."


async def _reply_forbidden(event: TelegramObject) -> None:
    if isinstance(event, Message):
        await event.answer(_FORBIDDEN_TEXT)
        return
    if isinstance(event, CallbackQuery):
        if event.message:
            await event.message.answer(_FORBIDDEN_TEXT)
        await event.answer("Недостаточно прав", show_alert=True)


class RoleMiddleware(BaseMiddleware):
    def __init__(self, *allowed_roles: str) -> None:
        if not allowed_roles:
            raise ValueError("RoleMiddleware requires at least one role")
        self._allowed_roles = allowed_roles

    async def __call__(
        self,
        handler: Callable[[TelegramObject, dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: dict[str, Any],
    ) -> Any:
        bot_user: BotUser | None = data.get("bot_user")
        if not has_role(bot_user, *self._allowed_roles):
            telegram_id = bot_user.telegram_id if bot_user else None
            role = bot_user.role if bot_user else None
            log_bot_security_event(
                "access_denied",
                telegram_id=telegram_id,
                role=role,
                action="role_middleware",
                detail=",".join(self._allowed_roles),
            )
            await _reply_forbidden(event)
            return None
        return await handler(event, data)
