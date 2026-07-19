"""Structured security audit logging for the Telegram bot."""

from __future__ import annotations

import logging
from typing import Any

_logger = logging.getLogger("bot.security")


def log_bot_security_event(
    event: str,
    *,
    telegram_id: int | None = None,
    role: str | None = None,
    action: str,
    detail: str | None = None,
    **extra: Any,
) -> None:
    parts = [
        f"event={event}",
        f"action={action}",
    ]
    if telegram_id is not None:
        parts.append(f"telegram_id={telegram_id}")
    if role is not None:
        parts.append(f"role={role}")
    if detail:
        parts.append(f"detail={detail}")
    for key, value in extra.items():
        if value is not None:
            parts.append(f"{key}={value}")
    _logger.info(" ".join(parts))
