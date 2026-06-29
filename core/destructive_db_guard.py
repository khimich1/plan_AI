"""Guards for irreversible SQLite reset operations (``clear_all_*``)."""

from __future__ import annotations

import logging
import os

logger = logging.getLogger(__name__)

_TRUTHY = frozenset({"1", "true", "yes", "on"})


class DestructiveDbOperationBlocked(RuntimeError):
    """Raised when ``clear_all_plates_data`` / ``clear_all_kp`` are not allowed."""

    def __init__(self, message: str | None = None) -> None:
        super().__init__(
            message
            or (
                "Операция полного обнуления БД запрещена. "
                "В development требуется ALLOW_DESTRUCTIVE_DB_RESET=1; "
                "в staging/production дополнительно DESTRUCTIVE_DB_RESET_BREAK_GLASS=1."
            )
        )


class DestructiveDbFlagsMisconfigured(RuntimeError):
    """Raised at startup when production is configured with full break-glass armed."""


def _env_truthy(name: str) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    return raw in _TRUTHY


def destructive_db_reset_allowed() -> bool:
    """True when destructive reset is permitted for the current environment.

    Fail-closed everywhere: ``development`` requires ``ALLOW_DESTRUCTIVE_DB_RESET``.
    ``production``, ``staging``, and any other non-development environment also
    require ``DESTRUCTIVE_DB_RESET_BREAK_GLASS`` (``ALLOW`` alone is not enough).
    """
    app_env = os.environ.get("APP_ENV", "development").strip().lower()
    allow = _env_truthy("ALLOW_DESTRUCTIVE_DB_RESET")
    if app_env == "development":
        return allow
    return allow and _env_truthy("DESTRUCTIVE_DB_RESET_BREAK_GLASS")


def require_destructive_db_reset() -> None:
    """Raise :class:`DestructiveDbOperationBlocked` if reset is not allowed."""
    if not destructive_db_reset_allowed():
        raise DestructiveDbOperationBlocked()


def fail_fast_if_destructive_flags_in_production(app_env: str) -> None:
    """Refuse production startup when both destructive flags are enabled."""
    if app_env.strip().lower() != "production":
        return

    allow = _env_truthy("ALLOW_DESTRUCTIVE_DB_RESET")
    break_glass = _env_truthy("DESTRUCTIVE_DB_RESET_BREAK_GLASS")
    if not allow and not break_glass:
        return

    if allow and break_glass:
        logger.critical(
            "APP_ENV=production with ALLOW_DESTRUCTIVE_DB_RESET and "
            "DESTRUCTIVE_DB_RESET_BREAK_GLASS — refusing to start."
        )
        raise DestructiveDbFlagsMisconfigured(
            "Destructive DB reset flags must not be fully armed in production."
        )

    if allow:
        logger.error(
            "APP_ENV=production with ALLOW_DESTRUCTIVE_DB_RESET set "
            "(DESTRUCTIVE_DB_RESET_BREAK_GLASS missing)."
        )
    if break_glass:
        logger.error(
            "APP_ENV=production with DESTRUCTIVE_DB_RESET_BREAK_GLASS set "
            "without ALLOW_DESTRUCTIVE_DB_RESET."
        )
