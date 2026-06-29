from __future__ import annotations

from typing import Any

from app.repositories.auth_repository import AuthRepository


class InvalidCurrentPasswordError(Exception):
    """Raised when the supplied current password does not match."""


class UserInactiveError(Exception):
    """Raised when the user account is inactive after a password change."""


class AuthService:
    def __init__(self, repository: AuthRepository | None = None) -> None:
        self.repository = repository or AuthRepository()

    def authenticate(self, username: str, password: str) -> dict[str, Any] | None:
        return self.repository.authenticate(username, password)

    def register(
        self,
        *,
        username: str,
        password: str,
        role: str,
        manager_id: int | None = None,
        is_active: bool = True,
    ) -> tuple[dict[str, Any], bool]:
        return self.repository.create_or_update_user(
            username=username,
            password=password,
            role=role,
            manager_id=manager_id,
            is_active=is_active,
        )

    def change_password(
        self,
        *,
        user: dict[str, Any],
        current_password: str,
        new_password: str,
    ) -> dict[str, Any]:
        authenticated = self.repository.authenticate(user["username"], current_password)
        if not authenticated:
            raise InvalidCurrentPasswordError()

        updated_user, _created = self.repository.create_or_update_user(
            username=user["username"],
            password=new_password,
            role=user["role"],
            manager_id=user.get("manager_id"),
            is_active=bool(user.get("is_active", True)),
        )
        refreshed_user = self.repository.get_user_by_id(int(user["id"]))
        if refreshed_user is None:
            raise UserInactiveError()

        new_session_version = self.repository.bump_session_version(int(user["id"]))
        refreshed_user["session_version"] = new_session_version
        return {"user": updated_user, "session_user": refreshed_user}
