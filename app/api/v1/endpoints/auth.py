from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, Request, Response, status

from app.dependencies.auth import get_auth_repository, get_current_user, require_roles
from app.dependencies.services import get_auth_service
from app.repositories.auth_repository import AuthRepository
from app.schemas.auth import (
    ChangePasswordRequest,
    ChangePasswordResponse,
    LoginRequest,
    LoginResponse,
    LogoutResponse,
    MeResponse,
    RegisterUserRequest,
    RegisterUserResponse,
)
from app.security.csrf import clear_csrf_cookie, generate_csrf_token, set_csrf_cookie
from app.security.login_rate_limit import (
    check_login_rate_limit,
    check_password_change_rate_limit,
    resolve_client_ip,
)
from app.security.password_policy import PasswordPolicyError
from app.security.session import (
    clear_session_cookie,
    create_session_token,
    decode_session_token,
    session_claims_from_user,
    set_session_cookie,
)
from app.services.auth_service import AuthService, InvalidCurrentPasswordError, UserInactiveError

router = APIRouter(prefix="/auth", tags=["auth"])


@router.post("/login", response_model=LoginResponse)
def login(
    payload: LoginRequest,
    request: Request,
    response: Response,
    auth_service: AuthService = Depends(get_auth_service),
) -> LoginResponse:
    check_login_rate_limit(resolve_client_ip(request))
    user = auth_service.authenticate(payload.username, payload.password)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid credentials")
    token = create_session_token(session_claims_from_user(user))
    set_session_cookie(response, token)
    set_csrf_cookie(response, generate_csrf_token())
    return LoginResponse.model_validate({"user": user})


@router.post("/logout", response_model=LogoutResponse)
def logout(
    request: Request,
    response: Response,
    repository: AuthRepository = Depends(get_auth_repository),
) -> LogoutResponse:
    token = request.cookies.get("app_session")
    payload = decode_session_token(token or "")
    if payload and payload.get("id") is not None:
        try:
            repository.bump_session_version(int(payload["id"]))
        except ValueError:
            pass
    clear_session_cookie(response)
    clear_csrf_cookie(response)
    return LogoutResponse(ok=True)


@router.get("/me", response_model=MeResponse)
def me(user: dict = Depends(get_current_user)) -> MeResponse:
    return MeResponse.model_validate({"user": user})


@router.post("/register", status_code=status.HTTP_201_CREATED, response_model=RegisterUserResponse)
def register_user(
    payload: RegisterUserRequest,
    _admin: dict = Depends(require_roles("admin")),
    auth_service: AuthService = Depends(get_auth_service),
) -> RegisterUserResponse:
    try:
        user, created = auth_service.register(
            username=payload.username,
            password=payload.password,
            role=payload.role,
            manager_id=payload.manager_id,
            is_active=payload.is_active,
        )
    except (ValueError, PasswordPolicyError) as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    return RegisterUserResponse.model_validate({"user": user, "created": created})


@router.post("/change-password", response_model=ChangePasswordResponse)
def change_password(
    payload: ChangePasswordRequest,
    request: Request,
    response: Response,
    user: dict = Depends(get_current_user),
    auth_service: AuthService = Depends(get_auth_service),
) -> ChangePasswordResponse:
    check_password_change_rate_limit(resolve_client_ip(request), int(user["id"]))
    try:
        result = auth_service.change_password(
            user=user,
            current_password=payload.current_password,
            new_password=payload.new_password,
        )
    except InvalidCurrentPasswordError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Current password is incorrect.",
        ) from exc
    except UserInactiveError as exc:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="User inactive") from exc
    except (ValueError, PasswordPolicyError) as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    token = create_session_token(session_claims_from_user(result["session_user"]))
    set_session_cookie(response, token)
    return ChangePasswordResponse.model_validate({"user": result["user"], "ok": True})
