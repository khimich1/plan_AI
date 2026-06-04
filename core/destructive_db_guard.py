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
                "Операция полного обнуления БД запрещена в production. "
                "Установите ALLOW_DESTRUCTIVE_DB_RESET=1 только для осознанного сброса."
            )
        )


def destructive_db_reset_allowed() -> bool:
    """True when destructive reset is permitted for the current environment."""
    app_env = os.environ.get("APP_ENV", "development").strip().lower()
    if app_env != "production":
        return True
    raw = os.environ.get("ALLOW_DESTRUCTIVE_DB_RESET", "").strip().lower()
    return raw in _TRUTHY


def require_destructive_db_reset() -> None:
    """Raise :class:`DestructiveDbOperationBlocked` if reset is not allowed."""
    if not destructive_db_reset_allowed():
        raise DestructiveDbOperationBlocked()
