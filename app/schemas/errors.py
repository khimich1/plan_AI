from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

ERROR_CODE_REST_VALIDATION_FAILED = "rest_validation_failed"
ERROR_CODE_UNPRICED_PLATES = "unpriced_plates"
ERROR_CODE_PLAN_VERSION_CONFLICT = "plan_version_conflict"

ApiErrorCode = Literal[
    "rest_validation_failed",
    "unpriced_plates",
    "plan_version_conflict",
]


class ApiErrorBody(BaseModel):
    code: ApiErrorCode | str = Field(description="Machine-readable error code")
    message: str = Field(description="Human-readable error message")
    details: dict | None = Field(default=None, description="Optional structured payload")
