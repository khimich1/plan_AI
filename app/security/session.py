from __future__ import annotations

import base64
import hashlib
import hmac
import json
import time
from typing import Any

from app.core.settings import get_settings


def _sign(payload: bytes) -> str:
    secret = get_settings().app_secret_key.encode("utf-8")
    digest = hmac.new(secret, payload, hashlib.sha256).digest()
    return base64.urlsafe_b64encode(digest).decode("ascii")


def create_session_token(data: dict[str, Any], ttl_seconds: int = 60 * 60 * 12) -> str:
    payload = dict(data)
    payload["exp"] = int(time.time()) + ttl_seconds
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

