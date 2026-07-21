from __future__ import annotations

import re

MIN_PASSWORD_LENGTH = 12

PASSWORD_TOO_SHORT_MESSAGE = (
    f"Password must be at least {MIN_PASSWORD_LENGTH} characters long."
)
PASSWORD_MISSING_UPPERCASE_MESSAGE = (
    "Password must contain at least one uppercase letter."
)
PASSWORD_MISSING_LOWERCASE_MESSAGE = (
    "Password must contain at least one lowercase letter."
)
PASSWORD_MISSING_DIGIT_MESSAGE = "Password must contain at least one digit."
PASSWORD_TOO_COMMON_MESSAGE = "Password is too common. Choose a more unique password."

_COMMON_PASSWORDS: frozenset[str] = frozenset(
    {
        "password",
        "password123",
        "password1234",
        "qwerty",
        "qwerty123",
        "qwertyuiop",
        "123456",
        "12345678",
        "123456789",
        "1234567890",
        "admin",
        "admin123",
        "administrator",
        "letmein",
        "welcome",
        "iloveyou",
        "monkey",
        "dragon",
        "master",
        "changeme",
        "football",
        "baseball",
        "sunshine",
        "princess",
        "login",
        "passw0rd",
        "trustno1",
    }
)

_HAS_UPPERCASE = re.compile(r"[A-Z]")
_HAS_LOWERCASE = re.compile(r"[a-z]")
_HAS_DIGIT = re.compile(r"\d")


class PasswordPolicyError(ValueError):
    """Raised when a password does not meet policy requirements."""


def validate_password(password: str) -> None:
    if len(password) < MIN_PASSWORD_LENGTH:
        raise PasswordPolicyError(PASSWORD_TOO_SHORT_MESSAGE)
    if password.casefold() in _COMMON_PASSWORDS:
        raise PasswordPolicyError(PASSWORD_TOO_COMMON_MESSAGE)
    if not _HAS_UPPERCASE.search(password):
        raise PasswordPolicyError(PASSWORD_MISSING_UPPERCASE_MESSAGE)
    if not _HAS_LOWERCASE.search(password):
        raise PasswordPolicyError(PASSWORD_MISSING_LOWERCASE_MESSAGE)
    if not _HAS_DIGIT.search(password):
        raise PasswordPolicyError(PASSWORD_MISSING_DIGIT_MESSAGE)
