from __future__ import annotations

import base64
import hashlib
import hmac
import json
import time
from typing import Any, Literal

from fastapi import Response

from app.core.settings import get_settings

SESSION_COOKIE_NAME = "app_session"

# Stateless HMAC cookies with server-side session_version (``sv`` claim in payload).
# - ``exp``: absolute expiry (idle/max age = SESSION_COOKIE_MAX_AGE at issuance).
# - ``sv``: must match ``app_users.session_version``; bumped on logout and password change.
# Rotating APP_SECRET_KEY still invalidates all sessions immediately.
# Full per-device revocation without global bump needs a sid revocation store — future work.


def _sign(payload: bytes) -> str:
    secret = get_settings().app_secret_key.encode("utf-8")
    digest = hmac.new(secret, payload, hashlib.sha256).digest()
    return base64.urlsafe_b64encode(digest).decode("ascii")


def create_session_token(
    data: dict[str, Any],
    ttl_seconds: int | None = None,
) -> str:
    settings = get_settings()
    ttl = ttl_seconds if ttl_seconds is not None else settings.session_ttl_seconds
    payload = dict(data)
    payload["exp"] = int(time.time()) + ttl
    encoded_payload = base64.urlsafe_b64encode(
        json.dumps(payload, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    signature = _sign(encoded_payload.encode("utf-8"))
    return f"{encoded_payload}.{signature}"


def decode_session_token(token: str) -> dict[str, Any] | None:
    if not token or "." not in token:
        return None
    encoded_payload, signature = token.rsplit(".", 1)
    expected = _sign(encoded_payload.encode("utf-8"))
    if not hmac.compare_digest(signature, expected):
        return None
    try:
        payload = json.loads(base64.urlsafe_b64decode(encoded_payload.encode("ascii")).decode("utf-8"))
    except Exception:
        return None
    if int(payload.get("exp", 0)) < int(time.time()):
        return None
    return payload


def token_session_version(payload: dict[str, Any]) -> int:
    """Return session version from token payload (legacy tokens without ``sv`` → 0)."""
    raw = payload.get("sv")
    if raw is None:
        return 0
    return int(raw)


def session_claims_from_user(user: dict[str, Any]) -> dict[str, Any]:
    """Build signed session payload fields from an authenticated user row."""
    return {
        "id": user["id"],
        "username": user["username"],
        "role": user["role"],
        "sv": int(user.get("session_version", 0)),
    }


def is_session_active(payload: dict[str, Any], user: dict[str, Any]) -> bool:
    """True when token session version matches the user's current server-side version."""
    return token_session_version(payload) == int(user.get("session_version", 0))


def session_cookie_policy() -> dict[str, bool | int | Literal["lax", "strict", "none"]]:
    """Shared attributes for set_cookie / delete_cookie (must match on logout)."""
    settings = get_settings()
    return {
        "httponly": True,
        "samesite": settings.cookie_samesite,
        "secure": settings.cookie_secure_enabled,
        "max_age": settings.session_ttl_seconds,
    }


def set_session_cookie(response: Response, token: str) -> None:
    policy = session_cookie_policy()
    response.set_cookie(
        SESSION_COOKIE_NAME,
        token,
        httponly=bool(policy["httponly"]),
        samesite=policy["samesite"],  # type: ignore[arg-type]
        secure=bool(policy["secure"]),
        max_age=int(policy["max_age"]),
    )


def clear_session_cookie(response: Response) -> None:
    policy = session_cookie_policy()
    response.delete_cookie(
        SESSION_COOKIE_NAME,
        httponly=bool(policy["httponly"]),
        samesite=policy["samesite"],  # type: ignore[arg-type]
        secure=bool(policy["secure"]),
    )
