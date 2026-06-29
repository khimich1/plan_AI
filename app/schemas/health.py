from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class RateLimitDeploymentInfo(BaseModel):
    model_config = ConfigDict(extra="allow")

    store: str
    shared_across_workers: bool
    configured_workers: int | None = None
    single_worker_required: bool
    deployment_note: str | None = None
    warning: str | None = None


class HealthResponse(BaseModel):
    status: str
    app: str | None = None
    environment: str | None = None
    rate_limiting: RateLimitDeploymentInfo | None = None
