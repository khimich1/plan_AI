from pathlib import Path

import pytest
from pydantic import ValidationError

from app.repositories.auth_repository import AuthRepository
from app.schemas.auth import ChangePasswordRequest, RegisterUserRequest
from app.security.password_policy import (
    MIN_PASSWORD_LENGTH,
    PASSWORD_MISSING_DIGIT_MESSAGE,
    PASSWORD_MISSING_LOWERCASE_MESSAGE,
    PASSWORD_MISSING_UPPERCASE_MESSAGE,
    PASSWORD_TOO_COMMON_MESSAGE,
    PASSWORD_TOO_SHORT_MESSAGE,
    PasswordPolicyError,
    validate_password,
)

_VALID_PASSWORD = "ValidPass1234"


def test_validate_password_accepts_strong_password() -> None:
    validate_password(_VALID_PASSWORD)


def test_validate_password_rejects_short_password() -> None:
    with pytest.raises(PasswordPolicyError, match=str(MIN_PASSWORD_LENGTH)):
        validate_password("Short1A")


def test_validate_password_rejects_missing_uppercase() -> None:
    with pytest.raises(PasswordPolicyError, match=PASSWORD_MISSING_UPPERCASE_MESSAGE):
        validate_password("validpass1234")


def test_validate_password_rejects_missing_lowercase() -> None:
    with pytest.raises(PasswordPolicyError, match=PASSWORD_MISSING_LOWERCASE_MESSAGE):
        validate_password("VALIDPASS1234")


def test_validate_password_rejects_missing_digit() -> None:
    with pytest.raises(PasswordPolicyError, match=PASSWORD_MISSING_DIGIT_MESSAGE):
        validate_password("ValidPassword")


def test_validate_password_rejects_common_password() -> None:
    with pytest.raises(PasswordPolicyError, match=PASSWORD_TOO_COMMON_MESSAGE):
        validate_password("password1234")


def test_register_user_request_rejects_weak_password() -> None:
    with pytest.raises(ValidationError):
        RegisterUserRequest(username="new_user", password="short", role="manager")


def test_register_user_request_accepts_strong_password() -> None:
    payload = RegisterUserRequest(
        username="new_user",
        password=_VALID_PASSWORD,
        role="manager",
    )
    assert payload.password == _VALID_PASSWORD


@pytest.mark.parametrize(
    "role",
    ["admin", "manager", "production", "logistics", "accountant"],
)
def test_register_user_request_accepts_valid_roles(role: str) -> None:
    payload = RegisterUserRequest(
        username="new_user",
        password=_VALID_PASSWORD,
        role=role,  # type: ignore[arg-type]
    )
    assert payload.role == role
    assert payload.password == _VALID_PASSWORD


def test_register_user_request_rejects_invalid_role() -> None:
    with pytest.raises(ValidationError):
        RegisterUserRequest(
            username="new_user",
            password=_VALID_PASSWORD,
            role="superadmin",
        )


def test_change_password_request_rejects_same_password() -> None:
    with pytest.raises(ValidationError, match="must differ"):
        ChangePasswordRequest(
            current_password=_VALID_PASSWORD,
            new_password=_VALID_PASSWORD,
        )


def test_create_or_update_user_rejects_short_password(tmp_path: Path) -> None:
    repository = AuthRepository(str(tmp_path / "auth.db"))

    with pytest.raises(PasswordPolicyError, match=PASSWORD_TOO_SHORT_MESSAGE):
        repository.create_or_update_user(
            username="new_user",
            password="short",
            role="manager",
        )


def test_create_or_update_user_rejects_common_password(tmp_path: Path) -> None:
    repository = AuthRepository(str(tmp_path / "auth.db"))

    with pytest.raises(PasswordPolicyError, match=PASSWORD_TOO_COMMON_MESSAGE):
        repository.create_or_update_user(
            username="new_user",
            password="password1234",
            role="manager",
        )
