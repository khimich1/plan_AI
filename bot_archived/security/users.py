"""Telegram bot user model and allowlist resolver."""

from __future__ import annotations

from dataclasses import dataclass

from core.config.settings import get_settings


@dataclass(frozen=True, slots=True)
class BotUser:
    telegram_id: int
    role: str
    app_user_id: int | None = None
    manager_id: int | None = None


def resolve_bot_user_from_db(telegram_id: int) -> BotUser | None:
    """DB lookup hook (BOTAUTH-010); no database access in MVP."""
    return None


def resolve_bot_user(telegram_id: int) -> BotUser | None:
    db_user = resolve_bot_user_from_db(telegram_id)
    if db_user is not None:
        return db_user
    settings = get_settings()
    role = settings.bot_telegram_allowlist.get(telegram_id)
    if role is None:
        return None
    return BotUser(telegram_id=telegram_id, role=role)


def has_role(user: BotUser | None, *roles: str) -> bool:
    if user is None:
        return False
    return user.role in roles
