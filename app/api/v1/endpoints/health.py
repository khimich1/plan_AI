from __future__ import annotations

from fastapi import APIRouter

from app.core.settings import get_settings
from app.security.login_rate_limit import rate_limit_deployment_info

router = APIRouter(tags=["health"])


@router.get("/health")
def healthcheck() -> dict:
    settings = get_settings()
    return {
        "status": "ok",
        "app": settings.app_name,
        "environment": settings.app_env,
        "rate_limiting": rate_limit_deployment_info(),
    }

