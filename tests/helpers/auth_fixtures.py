"""Auth test helpers — mock indexed user lookup for ``get_current_user``."""

from __future__ import annotations

from typing import Any

from app.repositories.auth_repository import AuthRepository


def patch_auth_users(monkeypatch: Any, users: list[dict[str, Any]]) -> None:
    """Patch ``AuthRepository.get_user_by_id`` for session auth in API tests."""
    by_id = {int(user["id"]): dict(user) for user in users}

    def get_user_by_id(self: AuthRepository, user_id: int) -> dict[str, Any] | None:
        return by_id.get(int(user_id))

    monkeypatch.setattr(AuthRepository, "get_user_by_id", get_user_by_id)
