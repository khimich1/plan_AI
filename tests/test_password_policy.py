from pathlib import Path

import pytest

from app.repositories.auth_repository import AuthRepository
from app.security.password_policy import (
    MIN_PASSWORD_LENGTH,
    PasswordPolicyError,
    validate_password,
)


def test_validate_password_accepts_min_length() -> None:
    validate_password("a" * MIN_PASSWORD_LENGTH)


def test_validate_password_rejects_short_password() -> None:
    with pytest.raises(PasswordPolicyError, match=str(MIN_PASSWORD_LENGTH)):
        validate_password("short")


def test_create_or_update_user_rejects_short_password(tmp_path: Path) -> None:
    repository = AuthRepository(str(tmp_path / "auth.db"))

    with pytest.raises(PasswordPolicyError):
        repository.create_or_update_user(
            username="new_user",
            password="short",
            role="manager",
        )
