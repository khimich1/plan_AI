from __future__ import annotations

import pytest
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient

from app.core.http_errors import (
    MSG_ARCHIVE_NOT_FOUND,
    raise_not_found_client_error,
    raise_structured_error,
    raise_unprocessable_client_error,
)
from app.schemas.errors import (
    ERROR_CODE_PLAN_VERSION_CONFLICT,
    ERROR_CODE_REST_VALIDATION_FAILED,
    ERROR_CODE_UNPRICED_PLATES,
    ApiErrorBody,
)


def test_api_error_body_model() -> None:
    body = ApiErrorBody(
        code=ERROR_CODE_UNPRICED_PLATES,
        message="Нет цен для части позиций",
        details={"positions": ["ПБ 78-12-8п"]},
    )
    assert body.model_dump(exclude_none=True) == {
        "code": "unpriced_plates",
        "message": "Нет цен для части позиций",
        "details": {"positions": ["ПБ 78-12-8п"]},
    }


def test_raise_structured_error_sets_http_exception_detail() -> None:
    with pytest.raises(HTTPException) as exc_info:
        raise_structured_error(
            status_code=422,
            code=ERROR_CODE_REST_VALIDATION_FAILED,
            message="Невалидный kp_id",
            details={"plan_id": "plan-1"},
            where="test_raise_structured_error",
        )

    exc = exc_info.value
    assert exc.status_code == 422
    assert exc.detail == {
        "code": "rest_validation_failed",
        "message": "Невалидный kp_id",
        "details": {"plan_id": "plan-1"},
    }


def test_raise_structured_error_omits_none_details() -> None:
    with pytest.raises(HTTPException) as exc_info:
        raise_structured_error(
            status_code=409,
            code=ERROR_CODE_PLAN_VERSION_CONFLICT,
            message="План был изменён другим запросом",
            where="test_raise_structured_error",
        )

    assert exc_info.value.detail == {
        "code": "plan_version_conflict",
        "message": "План был изменён другим запросом",
    }


def test_structured_error_json_response_shape() -> None:
    app = FastAPI()

    @app.get("/structured-error")
    def _trigger_structured_error() -> None:
        raise_structured_error(
            status_code=422,
            code=ERROR_CODE_UNPRICED_PLATES,
            message="Нет цен",
            details={"positions": ["ПБ 1"]},
            where="test_endpoint",
        )

    response = TestClient(app).get("/structured-error")
    assert response.status_code == 422
    assert response.json() == {
        "detail": {
            "code": "unpriced_plates",
            "message": "Нет цен",
            "details": {"positions": ["ПБ 1"]},
        }
    }


def test_raise_not_found_client_error_uses_safe_detail() -> None:
    with pytest.raises(HTTPException) as exc_info:
        raise_not_found_client_error(
            ValueError("internal kp_id=42 path=/secret"),
            where="test_not_found",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )

    assert exc_info.value.status_code == 404
    assert exc_info.value.detail == MSG_ARCHIVE_NOT_FOUND
    assert "secret" not in str(exc_info.value.detail)


def test_raise_unprocessable_client_error_hides_exception_text() -> None:
    with pytest.raises(HTTPException) as exc_info:
        raise_unprocessable_client_error(
            RuntimeError("sqlite3.OperationalError: no such table"),
            where="test_unprocessable",
        )

    assert exc_info.value.status_code == 422
    assert exc_info.value.detail == "Не удалось выполнить операцию. Проверьте введённые данные."
    assert "sqlite" not in str(exc_info.value.detail).lower()
