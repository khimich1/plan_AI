from __future__ import annotations

from typing import Annotated

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
    role: str = Field(min_length=1, max_length=32)
    manager_id: int | None = None
    is_active: bool = True


class ChangePasswordRequest(BaseModel):
    current_password: str = Field(min_length=1)
    new_password: StrongPassword

    @model_validator(mode="after")
    def new_password_must_differ(self) -> ChangePasswordRequest:
        if self.current_password == self.new_password:
            raise ValueError("New password must differ from the current password.")
        return self
