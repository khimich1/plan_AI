from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, Request, Response

from app.dependencies.auth import get_current_user
from app.repositories.auth_repository import AuthRepository
from app.schemas.auth import LoginRequest
from app.security.login_rate_limit import check_login_rate_limit, resolve_client_ip
from app.security.session import clear_session_cookie, create_session_token, set_session_cookie

router = APIRouter(prefix="/auth", tags=["auth"])


@router.post("/login")
def login(payload: LoginRequest, request: Request, response: Response) -> dict:
    check_login_rate_limit(resolve_client_ip(request))
    repository = AuthRepository()
    user = repository.authenticate(payload.username, payload.password)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid credentials")
    token = create_session_token(
        {
            "id": user["id"],
            "username": user["username"],
            "role": user["role"],
        }
    )
    set_session_cookie(response, token)
    return {"user": user}


@router.post("/logout")
def logout(response: Response) -> dict:
    clear_session_cookie(response)
    return {"ok": True}


@router.get("/me")
def me(user: dict = Depends(get_current_user)) -> dict:
    return {"user": user}

