from __future__ import annotations

from fastapi import APIRouter

from app.core.settings import get_settings
from app.schemas.health import HealthResponse
from app.security.login_rate_limit import rate_limit_deployment_info

router = APIRouter(tags=["health"])


def build_health_payload() -> dict:
    """Public health payload: redact deployment metadata in production (S9-audit)."""
    settings = get_settings()
    payload: dict = {"status": "ok"}
    if settings.app_env.lower() != "production":
        payload["app"] = settings.app_name
        payload["environment"] = settings.app_env
        payload["rate_limiting"] = rate_limit_deployment_info()
    return payload


@router.get("/health", response_model=HealthResponse, response_model_exclude_none=True)
def healthcheck() -> HealthResponse:
    return HealthResponse.model_validate(build_health_payload())
