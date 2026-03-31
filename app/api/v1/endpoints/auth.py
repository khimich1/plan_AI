from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, Response

from app.dependencies.auth import get_current_user
from app.repositories.auth_repository import AuthRepository
from app.schemas.auth import LoginRequest
from app.security.session import create_session_token

router = APIRouter(prefix="/auth", tags=["auth"])


@router.post("/login")
def login(payload: LoginRequest, response: Response) -> dict:
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
    response.set_cookie(
        "app_session",
        token,
        httponly=True,
        samesite="lax",
        secure=False,
        max_age=60 * 60 * 12,
    )
    return {"user": user}


@router.post("/logout")
def logout(response: Response) -> dict:
    response.delete_cookie("app_session")
    return {"ok": True}


@router.get("/me")
def me(user: dict = Depends(get_current_user)) -> dict:
    return {"user": user}

