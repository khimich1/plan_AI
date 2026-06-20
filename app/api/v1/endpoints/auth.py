from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, Request, Response, status

from app.dependencies.auth import get_auth_repository, get_current_user, require_roles
from app.repositories.auth_repository import AuthRepository
from app.schemas.auth import ChangePasswordRequest, LoginRequest, RegisterUserRequest
from app.security.csrf import clear_csrf_cookie, generate_csrf_token, set_csrf_cookie
from app.security.login_rate_limit import check_login_rate_limit, resolve_client_ip
from app.security.password_policy import PasswordPolicyError
from app.security.session import (
    clear_session_cookie,
    create_session_token,
    decode_session_token,
    session_claims_from_user,
    set_session_cookie,
)

router = APIRouter(prefix="/auth", tags=["auth"])


@router.post("/login")
def login(payload: LoginRequest, request: Request, response: Response) -> dict:
    check_login_rate_limit(resolve_client_ip(request))
    repository = AuthRepository()
    user = repository.authenticate(payload.username, payload.password)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid credentials")
    token = create_session_token(session_claims_from_user(user))
    set_session_cookie(response, token)
    set_csrf_cookie(response, generate_csrf_token())
    return {"user": user}


@router.post("/logout")
def logout(
    request: Request,
    response: Response,
    repository: AuthRepository = Depends(get_auth_repository),
) -> dict:
    token = request.cookies.get("app_session")
    payload = decode_session_token(token or "")
    if payload and payload.get("id") is not None:
        try:
            repository.bump_session_version(int(payload["id"]))
        except ValueError:
            pass
    clear_session_cookie(response)
    clear_csrf_cookie(response)
    return {"ok": True}


@router.get("/me")
def me(user: dict = Depends(get_current_user)) -> dict:
    return {"user": user}


@router.post("/register", status_code=status.HTTP_201_CREATED)
def register_user(
    payload: RegisterUserRequest,
    _admin: dict = Depends(require_roles("admin")),
) -> dict:
    repository = AuthRepository()
    try:
        user, created = repository.create_or_update_user(
            username=payload.username,
            password=payload.password,
            role=payload.role,
            manager_id=payload.manager_id,
            is_active=payload.is_active,
        )
    except (ValueError, PasswordPolicyError) as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    return {"user": user, "created": created}


@router.post("/change-password")
def change_password(
    payload: ChangePasswordRequest,
    response: Response,
    user: dict = Depends(get_current_user),
    repository: AuthRepository = Depends(get_auth_repository),
) -> dict:
    authenticated = repository.authenticate(user["username"], payload.current_password)
    if not authenticated:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Current password is incorrect.",
        )
    try:
        updated_user, _created = repository.create_or_update_user(
            username=user["username"],
            password=payload.new_password,
            role=user["role"],
            manager_id=user.get("manager_id"),
            is_active=bool(user.get("is_active", True)),
        )
    except (ValueError, PasswordPolicyError) as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    refreshed_user = repository.get_user_by_id(int(user["id"]))
    if refreshed_user is None:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="User inactive")
    new_session_version = repository.bump_session_version(int(user["id"]))
    refreshed_user["session_version"] = new_session_version
    token = create_session_token(session_claims_from_user(refreshed_user))
    set_session_cookie(response, token)
    return {"user": updated_user, "ok": True}

