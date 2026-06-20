"""Guards for irreversible SQLite reset operations (``clear_all_*``)."""

from __future__ import annotations

import os

_TRUTHY = frozenset({"1", "true", "yes", "on"})


class DestructiveDbOperationBlocked(RuntimeError):
    """Raised when ``clear_all_plates_data`` / ``clear_all_kp`` are not allowed."""

    def __init__(self, message: str | None = None) -> None:
        super().__init__(
            message
            or (
                "Операция полного обнуления БД запрещена в production/staging. "
                "Требуются ALLOW_DESTRUCTIVE_DB_RESET=1 и DESTRUCTIVE_DB_RESET_BREAK_GLASS=1."
            )
        )


def _env_truthy(name: str) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    return raw in _TRUTHY


def destructive_db_reset_allowed() -> bool:
    """True when destructive reset is permitted for the current environment.

    Allowed in ``development`` without extra flags.
    In ``production``, ``staging``, and any other non-development environment,
    both ``ALLOW_DESTRUCTIVE_DB_RESET`` and ``DESTRUCTIVE_DB_RESET_BREAK_GLASS``
    must be set (fail-closed: ``ALLOW`` alone is not enough).
    """
    app_env = os.environ.get("APP_ENV", "development").strip().lower()
    if app_env == "development":
        return True
    return _env_truthy("ALLOW_DESTRUCTIVE_DB_RESET") and _env_truthy(
        "DESTRUCTIVE_DB_RESET_BREAK_GLASS"
    )


def require_destructive_db_reset() -> None:
    """Raise :class:`DestructiveDbOperationBlocked` if reset is not allowed."""
    if not destructive_db_reset_allowed():
        raise DestructiveDbOperationBlocked()
