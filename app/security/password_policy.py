from __future__ import annotations

MIN_PASSWORD_LENGTH = 8

PASSWORD_TOO_SHORT_MESSAGE = (
    f"Password must be at least {MIN_PASSWORD_LENGTH} characters long."
)


class PasswordPolicyError(ValueError):
    """Raised when a password does not meet policy requirements."""


def validate_password(password: str) -> None:
    if len(password) < MIN_PASSWORD_LENGTH:
        raise PasswordPolicyError(PASSWORD_TOO_SHORT_MESSAGE)
