from __future__ import annotations

from typing import Annotated, Literal

from pydantic import AfterValidator, BaseModel, Field, model_validator

from app.security.password_policy import PasswordPolicyError, validate_password


def _validate_strong_password(value: str) -> str:
    try:
        validate_password(value)
    except PasswordPolicyError as exc:
        raise ValueError(str(exc)) from exc
    return value


StrongPassword = Annotated[str, AfterValidator(_validate_strong_password)]


class LoginRequest(BaseModel):
    username: str = Field(min_length=1)
    password: str = Field(min_length=1)


class RegisterUserRequest(BaseModel):
    username: str = Field(min_length=1, max_length=64)
    password: StrongPassword
    role: Literal["admin", "manager", "production"]
    manager_id: int | None = None
    is_active: bool = True


class UserPublic(BaseModel):
    id: int
    username: str
    role: str
    manager_id: int | None = None
    is_active: int
    created_at: str


class AuthUserResponse(BaseModel):
    id: int
    username: str
    role: str
    manager_id: int | None = None
    is_active: int
    session_version: int
    created_at: str


class LoginResponse(BaseModel):
    user: AuthUserResponse


class MeResponse(BaseModel):
    user: AuthUserResponse


class LogoutResponse(BaseModel):
    ok: bool = True


class RegisterUserResponse(BaseModel):
    user: AuthUserResponse
    created: bool


class ChangePasswordResponse(BaseModel):
    user: AuthUserResponse
    ok: bool = True


class UsersPageResponse(BaseModel):
    items: list[UserPublic]
    total: int = Field(ge=0)
    limit: int = Field(ge=1)
    offset: int = Field(ge=0)


class ChangePasswordRequest(BaseModel):
    current_password: str = Field(min_length=1)
    new_password: StrongPassword

    @model_validator(mode="after")
    def new_password_must_differ(self) -> ChangePasswordRequest:
        if self.current_password == self.new_password:
            raise ValueError("New password must differ from the current password.")
        return self
