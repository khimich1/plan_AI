"""Per-IP login rate limiting (in-process sliding window).

Deployment constraint (audit S3 / P1-next WP4):
    Counters live in process memory and are **not** shared across uvicorn/gunicorn
    workers. With ``N`` workers, an attacker effectively gets ``N × limit`` attempts
    unless traffic is pinned to one worker.

    Production options (pick one):

    1. **Single worker (recommended):** ``uvicorn app.main:app --workers 1``
    2. **Sticky sessions:** load balancer session affinity by client IP (weaker —
       counters reset on worker restart; still no cross-worker sharing)
    3. **Shared store (future):** Redis or similar — not configured in this project

    Set ``UVICORN_WORKERS`` or ``WEB_CONCURRENCY`` in the environment so startup
    logs a warning when ``> 1`` without a shared store. ``GET /health`` exposes
    the same metadata for ops checks.

    OCR upload limits (``commercial_upload_validation``) follow the same model.

See: ``ai_docs/develop/architecture/rate-limiting.md``.
"""

from __future__ import annotations

import logging
import math
import os
import threading
import time
from collections import defaultdict
from typing import Any

from fastapi import HTTPException, Request, status

from app.core.settings import get_settings

logger = logging.getLogger(__name__)

MSG_LOGIN_RATE_LIMIT = "Слишком много попыток входа. Повторите позже."

_MULTI_WORKER_ENV_VARS = ("UVICORN_WORKERS", "WEB_CONCURRENCY")
_DEPLOYMENT_NOTE = (
    "Rate limits are in-process only. Use uvicorn --workers 1 or sticky sessions."
)


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


def configured_worker_count() -> int | None:
    """Return worker count from deployment env vars, if set."""
    for name in _MULTI_WORKER_ENV_VARS:
        raw = os.environ.get(name, "").strip()
        if raw.isdigit():
            return int(raw)
    return None


def rate_limit_deployment_info() -> dict[str, Any]:
    """Metadata for health checks and deployment verification."""
    workers = configured_worker_count()
    info: dict[str, Any] = {
        "store": "in-process",
        "shared_across_workers": False,
        "configured_workers": workers,
        "single_worker_required": True,
        "deployment_note": _DEPLOYMENT_NOTE,
    }
    if workers is not None and workers > 1:
        info["warning"] = (
            f"configured_workers={workers} without shared store; "
            "effective limits are split across workers"
        )
    return info


def warn_if_multi_worker_without_shared_store() -> None:
    """Log at startup when env declares multiple workers without a shared store."""
    workers = configured_worker_count()
    if workers is None or workers <= 1:
        return
    logger.warning(
        "Rate limiting uses in-process store only; %s means limits are not "
        "shared across workers. Use `uvicorn app.main:app --workers 1`, sticky "
        "sessions, or add a shared store. See "
        "ai_docs/develop/architecture/rate-limiting.md.",
        f"configured_workers={workers}",
    )


def resolve_client_ip(request: Request) -> str:
    """Resolve client IP for rate limiting (audit S7).

    ``X-Forwarded-For`` is used only when the direct TCP peer
    (``request.client.host``) is in ``TRUSTED_PROXY_IPS``. With an empty
    default, XFF is never trusted — prevents spoofing when the app is exposed
    without a reverse proxy.

    Behind nginx or a load balancer, set ``TRUSTED_PROXY_IPS`` to the proxy
    address(es) as the app sees them (e.g. ``127.0.0.1`` on the same host).
    """
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
