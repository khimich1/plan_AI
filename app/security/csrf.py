from __future__ import annotations

import secrets
from typing import Literal

from fastapi import Response

from app.core.settings import get_settings

CSRF_COOKIE_NAME = "csrf_token"
CSRF_HEADER_NAME = "X-CSRF-Token"
CSRF_FORM_FIELD = "csrf_token"

_SAFE_METHODS = frozenset({"GET", "HEAD", "OPTIONS", "TRACE"})


def generate_csrf_token() -> str:
    return secrets.token_urlsafe(32)


def csrf_cookie_policy() -> dict[str, bool | int | Literal["lax", "strict", "none"]]:
    """CSRF cookie is readable by JS (not HttpOnly) for double-submit header pattern."""
    settings = get_settings()
    return {
        "httponly": False,
        "samesite": settings.cookie_samesite,
        "secure": settings.cookie_secure_enabled,
        "max_age": settings.session_ttl_seconds,
    }


def set_csrf_cookie(response: Response, token: str) -> None:
    policy = csrf_cookie_policy()
    response.set_cookie(
        CSRF_COOKIE_NAME,
        token,
        httponly=bool(policy["httponly"]),
        samesite=policy["samesite"],  # type: ignore[arg-type]
        secure=bool(policy["secure"]),
        max_age=int(policy["max_age"]),
        path="/",
    )


def clear_csrf_cookie(response: Response) -> None:
    policy = csrf_cookie_policy()
    response.delete_cookie(
        CSRF_COOKIE_NAME,
        httponly=bool(policy["httponly"]),
        samesite=policy["samesite"],  # type: ignore[arg-type]
        secure=bool(policy["secure"]),
        path="/",
    )


def is_safe_method(method: str) -> bool:
    return method.upper() in _SAFE_METHODS


def tokens_match(cookie_token: str | None, submitted_token: str | None) -> bool:
    if not cookie_token or not submitted_token:
        return False
    # secrets.compare_digest requires same-length strings; normalize via strip only.
    cookie = cookie_token.strip()
    submitted = submitted_token.strip()
    if not cookie or not submitted or len(cookie) != len(submitted):
        return False
    return secrets.compare_digest(cookie, submitted)
