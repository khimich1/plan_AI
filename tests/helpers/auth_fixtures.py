"""Auth test helpers — mock indexed user lookup for ``get_current_user``."""

from __future__ import annotations

from typing import Any

from app.repositories.auth_repository import AuthRepository


def patch_auth_users(monkeypatch: Any, users: list[dict[str, Any]]) -> None:
    """Patch ``AuthRepository.get_user_by_id`` for session auth in API tests."""
    by_id = {int(user["id"]): dict(user) for user in users}

    def get_user_by_id(self: AuthRepository, user_id: int) -> dict[str, Any] | None:
        user = by_id.get(int(user_id))
        if user is None:
            return None
        payload = dict(user)
        payload.setdefault("session_version", 0)
        return payload

    monkeypatch.setattr(AuthRepository, "get_user_by_id", get_user_by_id)


def patch_auth_login(
    monkeypatch: Any,
    *,
    username: str = "admin",
    password: str = "StrongPassword123!",
    user_id: int = 1,
    role: str = "admin",
    session_version: int = 0,
) -> dict[str, Any]:
    """Patch login + session validation for a single test user."""
    user: dict[str, Any] = {
        "id": user_id,
        "username": username,
        "role": role,
        "session_version": session_version,
    }

    def fake_authenticate(self: AuthRepository, user_name: str, pwd: str) -> dict[str, Any] | None:
        if user_name == username and pwd == password:
            return dict(user)
        return None

    def fake_get_user_by_id(self: AuthRepository, lookup_id: int) -> dict[str, Any] | None:
        if int(lookup_id) != int(user["id"]):
            return None
        return {
            "id": user["id"],
            "username": user["username"],
            "role": user["role"],
            "manager_id": None,
            "is_active": 1,
            "session_version": user["session_version"],
            "created_at": "2026-01-01 00:00:00",
        }

    def fake_bump_session_version(self: AuthRepository, lookup_id: int) -> int:
        if int(lookup_id) != int(user["id"]):
            raise ValueError(f"User {lookup_id} not found.")
        user["session_version"] = int(user["session_version"]) + 1
        return int(user["session_version"])

    monkeypatch.setattr(AuthRepository, "authenticate", fake_authenticate)
    monkeypatch.setattr(AuthRepository, "get_user_by_id", fake_get_user_by_id)
    monkeypatch.setattr(AuthRepository, "bump_session_version", fake_bump_session_version)
    return user
