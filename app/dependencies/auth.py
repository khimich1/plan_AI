from __future__ import annotations

from fastapi import Depends, HTTPException, Request, status

from app.repositories.auth_repository import AuthRepository
from app.security.session import decode_session_token


def get_auth_repository() -> AuthRepository:
    return AuthRepository()


def get_current_user(
    request: Request,
    repository: AuthRepository = Depends(get_auth_repository),
) -> dict:
    token = request.cookies.get("app_session")
    payload = decode_session_token(token or "")
    if not payload:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Not authenticated")
    users = {user["id"]: user for user in repository.list_users()}
    user = users.get(int(payload["id"]))
    if not user or not user.get("is_active"):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="User inactive")
    return user


def require_roles(*allowed_roles: str):
    def dependency(user: dict = Depends(get_current_user)) -> dict:
        if user.get("role") not in allowed_roles:
            raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Forbidden")
        return user

    return dependency


# Shared dependency object so FastAPI caches auth once per request across multiple Depends(...).
REQUIRE_ADMIN_OR_MANAGER = require_roles("admin", "manager")

