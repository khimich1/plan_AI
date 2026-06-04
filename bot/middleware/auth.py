"""Global Telegram bot authentication middleware."""

from __future__ import annotations

from typing import Any, Awaitable, Callable

from aiogram import BaseMiddleware
from aiogram.types import CallbackQuery, Message, TelegramObject, User

from bot.security.audit import log_bot_security_event
from bot.security.users import BotUser, resolve_bot_user
from core.config.settings import get_settings

_ACCESS_DENIED_TEXT = (
    "⛔ Доступ запрещён.\n"
    "Ваш Telegram-аккаунт не в списке разрешённых пользователей.\n"
    "Обратитесь к администратору."
)


def _extract_user(event: TelegramObject) -> User | None:
    direct = getattr(event, "from_user", None)
    if direct is not None:
        return direct
    for attr in (
        "message",
        "edited_message",
        "callback_query",
        "inline_query",
        "chosen_inline_result",
        "shipping_query",
        "pre_checkout_query",
        "my_chat_member",
        "chat_member",
    ):
        inner = getattr(event, attr, None)
        if inner is None:
            continue
        user = getattr(inner, "from_user", None)
        if user is not None:
            return user
    return None


async def _reply_access_denied(event: TelegramObject) -> None:
    if isinstance(event, Message):
        await event.answer(_ACCESS_DENIED_TEXT)
        return
    if isinstance(event, CallbackQuery):
        if event.message:
            await event.message.answer(_ACCESS_DENIED_TEXT)
        await event.answer("Доступ запрещён", show_alert=True)
        return
    message = getattr(event, "message", None)
    if isinstance(message, Message):
        await message.answer(_ACCESS_DENIED_TEXT)


class BotAuthMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: dict[str, Any],
    ) -> Any:
        settings = get_settings()
        user = _extract_user(event)

        if not settings.bot_auth_enabled:
            if settings.app_env.lower() == "production":
                log_bot_security_event(
                    "misconfiguration",
                    action="bot_auth_disabled_in_production",
                )
                return None
            telegram_id = user.id if user else 0
            data["bot_user"] = BotUser(telegram_id=telegram_id, role="admin")
            return await handler(event, data)

        if user is None:
            log_bot_security_event(
                "access_denied",
                action="missing_telegram_user",
            )
            return None

        bot_user = resolve_bot_user(user.id)
        if bot_user is None:
            log_bot_security_event(
                "access_denied",
                telegram_id=user.id,
                action="allowlist_miss",
            )
            await _reply_access_denied(event)
            return None

        data["bot_user"] = bot_user
        return await handler(event, data)
