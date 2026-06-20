from __future__ import annotations

import math
import threading
import time
from collections import defaultdict

from fastapi import HTTPException, Request, status

from app.core.settings import get_settings

MSG_LOGIN_RATE_LIMIT = "Слишком много попыток входа. Повторите позже."


class _SlidingWindowRateLimiter:
    """In-process sliding window (not shared across workers)."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._events: dict[str, list[float]] = defaultdict(list)

    def check(self, key: str, *, max_events: int, window_seconds: float) -> None:
        now = time.monotonic()
        with self._lock:
            events = self._events[key]
            cutoff = now - window_seconds
            while events and events[0] < cutoff:
                events.pop(0)
            if len(events) >= max_events:
                retry_after = max(1, math.ceil(events[0] + window_seconds - now))
                raise HTTPException(
                    status_code=status.HTTP_429_TOO_MANY_REQUESTS,
                    detail=MSG_LOGIN_RATE_LIMIT,
                    headers={"Retry-After": str(retry_after)},
                )
            events.append(now)

    def reset(self) -> None:
        with self._lock:
            self._events.clear()


_login_rate_limiter = _SlidingWindowRateLimiter()


def reset_login_rate_limiter_for_tests() -> None:
    _login_rate_limiter.reset()


def resolve_client_ip(request: Request) -> str:
    direct_host = request.client.host if request.client and request.client.host else None
    if direct_host and direct_host in get_settings().trusted_proxy_ips:
        forwarded = request.headers.get("X-Forwarded-For")
        if forwarded:
            first = forwarded.split(",")[0].strip()
            if first:
                return first
    if direct_host:
        return direct_host
    return "unknown"


def check_login_rate_limit(client_ip: str) -> None:
    settings = get_settings()
    _login_rate_limiter.check(
        client_ip,
        max_events=settings.auth_login_attempts_per_minute,
        window_seconds=float(settings.auth_login_rate_limit_window_seconds),
    )
