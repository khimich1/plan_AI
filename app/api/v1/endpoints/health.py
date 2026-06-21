from __future__ import annotations

from fastapi import APIRouter

from app.core.settings import get_settings
from app.security.login_rate_limit import rate_limit_deployment_info

router = APIRouter(tags=["health"])


def build_health_payload() -> dict:
    """Public health payload: redact deployment metadata in production (S9-audit)."""
    settings = get_settings()
    payload: dict = {
        "status": "ok",
        "rate_limiting": rate_limit_deployment_info(),
    }
    if settings.app_env.lower() != "production":
        payload["app"] = settings.app_name
        payload["environment"] = settings.app_env
    return payload


@router.get("/health")
def healthcheck() -> dict:
    return build_health_payload()

