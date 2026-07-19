"""Промежуточные слои aiogram."""

from bot.middleware.auth import BotAuthMiddleware
from bot.middleware.role import RoleMiddleware

__all__ = ["BotAuthMiddleware", "RoleMiddleware"]
