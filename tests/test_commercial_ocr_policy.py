from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.services.commercial_upload_validation import (
    MSG_EXTERNAL_OCR_DISABLED,
    check_commercial_ocr_rate_limit,
    ensure_external_ocr_enabled,
    reset_commercial_ocr_rate_limiter_for_tests,
)
from tests.test_commercial_web_flow import _MINIMAL_PNG_BYTES

pytest_plugins = ["tests.test_commercial_web_flow"]


def test_ensure_external_ocr_enabled_raises_when_disabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from fastapi import HTTPException

    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "false")
    get_settings.cache_clear()
    with pytest.raises(HTTPException) as exc_info:
        ensure_external_ocr_enabled()
    assert exc_info.value.status_code == 503
    assert MSG_EXTERNAL_OCR_DISABLED in str(exc_info.value.detail)


def test_ensure_external_ocr_enabled_passes_when_enabled(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "true")
    get_settings.cache_clear()
    ensure_external_ocr_enabled()


def test_commercial_upload_blocked_when_ocr_disabled(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "false")
    get_settings.cache_clear()

    response = client.post(
        "/api/v1/commercial/from-form",
        data={
            "text": "",
            "manager_id": "1",
            "client_name": "ООО Тест",
            "discount_percent": "0",
            "delivery_conditions": "",
            "payment_conditions": "",
        },
        files={"image": ("a.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 503
    assert MSG_EXTERNAL_OCR_DISABLED in response.json()["detail"]


def test_commercial_text_only_works_when_ocr_disabled(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService
    from tests.test_commercial_web_flow import _sample_draft

    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "false")
    get_settings.cache_clear()

    async def fake_create(self, **kwargs):
        assert kwargs["image_bytes"] is None
        return _sample_draft("draft-text-only")

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", fake_create)

    response = client.post(
        "/api/v1/commercial/from-form",
        data={
            "text": "ПБ 78-12-8п 2",
            "manager_id": "1",
            "client_name": "ООО Тест",
            "discount_percent": "0",
            "delivery_conditions": "",
            "payment_conditions": "",
        },
    )
    assert response.status_code == 200


def test_commercial_parse_text_works_when_ocr_disabled(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.domain.models.parse_result import ParseResult
    from app.domain.models.plate_order import PlateOrder
    from app.services.commercial_service import CommercialService

    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "false")
    get_settings.cache_clear()

    order = PlateOrder()
    monkeypatch.setattr(
        CommercialService,
        "parse",
        lambda self, text: ParseResult(
            order=order,
            normalized_text=text,
            normalized_lines=[text],
            unparsed_lines=[],
            warnings=[],
            wide_plate_lines=[],
            diagnostics={},
        ),
    )

    response = client.post(
        "/api/v1/commercial/parse",
        json={"text": "ПБ 78-12-8п 2"},
    )
    assert response.status_code == 200


def test_apply_ai_plates_blocked_when_ocr_disabled(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    from app.domain.models.optimization_context import OptimizationContext
    from app.domain.models.plate_order import PlateOrder
    from app.services.draft_store import DraftStore

    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "false")
    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()

    draft_id = "e" * 32
    store = DraftStore()
    order = PlateOrder()
    oc = OptimizationContext(order=order)
    store.replace_preview(
        draft_id,
        order=order,
        optimization_context=oc,
        order_data=[],
        metadata={"owner_user_id": 1, "input_text": "ПБ 78-12-8п 2"},
    )

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/plates/ai",
        data={"instruction": "замени 78 на 60"},
    )
    assert response.status_code == 503
    assert MSG_EXTERNAL_OCR_DISABLED in response.json()["detail"]


def test_commercial_upload_allowed_when_ocr_enabled(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService
    from app.services.commercial_upload_validation import reset_commercial_ocr_rate_limiter_for_tests
    from tests.test_commercial_web_flow import _sample_draft

    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "true")
    get_settings.cache_clear()
    reset_commercial_ocr_rate_limiter_for_tests()

    async def fake_create(self, **kwargs):
        assert kwargs["image_bytes"] == _MINIMAL_PNG_BYTES
        return _sample_draft("draft-ocr-on")

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", fake_create)

    response = client.post(
        "/api/v1/commercial/from-form",
        data={
            "text": "",
            "manager_id": "1",
            "client_name": "ООО Тест",
            "discount_percent": "0",
            "delivery_conditions": "",
            "payment_conditions": "",
        },
        files={"image": ("a.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 200


def test_check_commercial_ocr_rate_limit_skips_when_zero(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("COMMERCIAL_OCR_UPLOADS_PER_HOUR", "0")
    get_settings.cache_clear()
    reset_commercial_ocr_rate_limiter_for_tests()
    for _ in range(11):
        check_commercial_ocr_rate_limit(1)
