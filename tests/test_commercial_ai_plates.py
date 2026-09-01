from __future__ import annotations

from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.services.commercial_workflow_service import CommercialWorkflowService
from core.plate_order_context import PlateOrderContext
from tests.test_commercial_web_flow import _MINIMAL_PNG_BYTES, _sample_draft

pytest_plugins = ["tests.test_commercial_web_flow"]


def test_apply_ai_plates_endpoint(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from app.domain.models.optimization_context import OptimizationContext
    from app.domain.models.plate_order import PlateOrder
    from app.services.draft_store import DraftStore

    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()

    draft_id = "c" * 32
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

    captured: dict = {}

    async def fake_apply(self, draft_id_arg: str, **kwargs):
        captured["draft_id"] = draft_id_arg
        captured.update(kwargs)
        draft = _sample_draft(draft_id_arg)
        draft["metadata"]["source_type"] = "ai"
        draft["metadata"]["ai_applied"] = True
        draft["metadata"]["last_ai_instruction"] = kwargs["instruction"]
        draft["metadata"]["input_text"] = "ПБ 60-12-8п 7"
        draft["metadata"]["normalized_text"] = "ПБ 60-12-8п 7"
        return draft

    monkeypatch.setattr(CommercialWorkflowService, "apply_ai_plates_instruction", fake_apply)

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/plates/ai",
        data={"instruction": "замени все 78 на 60"},
        files={"image": ("plates.png", _MINIMAL_PNG_BYTES, "image/png")},
    )
    assert response.status_code == 200
    body = response.json()
    assert body["metadata"]["source_type"] == "ai"
    assert body["metadata"]["ai_applied"] is True
    assert captured["instruction"] == "замени все 78 на 60"
    assert captured["image_bytes"] == _MINIMAL_PNG_BYTES


def test_apply_ai_plates_requires_instruction(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from app.domain.models.optimization_context import OptimizationContext
    from app.domain.models.plate_order import PlateOrder
    from app.services.draft_store import DraftStore

    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    get_settings.cache_clear()

    draft_id = "d" * 32
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
        data={"instruction": "  "},
    )
    assert response.status_code == 400


def test_apply_ai_plates_provider_error_returns_502(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from openai import PermissionDeniedError
    from app.domain.models.optimization_context import OptimizationContext
    from app.domain.models.plate_order import PlateOrder
    from app.services.draft_store import DraftStore
    from app.core.http_errors import MSG_AI_PROVIDER_UNAVAILABLE, MSG_INTERNAL

    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    monkeypatch.setenv("OCR_EXTERNAL_ENABLED", "true")
    get_settings.cache_clear()

    draft_id = "f" * 32
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

    async def fake_apply(self, draft_id_arg: str, **kwargs):
        request = MagicMock()
        response = MagicMock()
        response.status_code = 403
        response.headers = {}
        response.request = request
        raise PermissionDeniedError(
            "Country, region, or territory not supported",
            response=response,
            body={
                "error": {
                    "code": "unsupported_country_region_territory",
                    "message": "Country, region, or territory not supported",
                }
            },
        )

    monkeypatch.setattr(CommercialWorkflowService, "apply_ai_plates_instruction", fake_apply)

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/plates/ai",
        data={"instruction": "замени все 78 на 60"},
    )
    assert response.status_code == 502
    assert response.json()["detail"] == MSG_AI_PROVIDER_UNAVAILABLE
    assert MSG_INTERNAL not in response.json()["detail"]


async def _run_apply_ai_workflow(monkeypatch, tmp_path):
    from app.services.draft_store import DraftStore
    from app.domain.models.optimization_context import OptimizationContext
    from app.domain.models.plate_order import PlateOrder

    drafts_dir = tmp_path / "drafts"
    drafts_dir.mkdir()
    monkeypatch.setenv("DRAFTS_DIR", str(drafts_dir))
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    get_settings.cache_clear()

    draft_id = "b" * 32
    store = DraftStore()
    order = PlateOrder()
    oc = OptimizationContext(order=order)
    store.replace_preview(
        draft_id,
        order=order,
        optimization_context=oc,
        order_data=[],
        metadata={
            "owner_user_id": 1,
            "input_text": "ПБ 78-12-8п 2",
            "normalized_text": "ПБ 78-12-8п 2",
            "normalized_lines": ["ПБ 78-12-8п 2"],
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
    )

    async def fake_ai(**kwargs):
        return {
            "text": "ПБ 60-12-8п 7",
            "plates": [
                {
                    "raw_name": "ПБ 60-12-8п",
                    "normalized_candidate": "ПБ 60-12-8п",
                    "qty": 7,
                    "confidence": 0.99,
                    "issues": [],
                }
            ],
            "draft_plates": [],
            "corrections": [],
            "method": "GPT-4o+ai",
            "cost_usd": 0.001,
        }

    monkeypatch.setattr(
        "app.services.commercial_workflow_service.apply_plates_with_ai",
        AsyncMock(side_effect=fake_ai),
    )

    workflow = CommercialWorkflowService()
    result = await workflow.apply_ai_plates_instruction(
        draft_id,
        instruction="замени 78 на 60 и qty 7",
        image_bytes=None,
        image_filename=None,
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )
    assert result["metadata"]["input_text"] == "ПБ 60-12-8п 7"
    assert result["metadata"]["source_type"] == "ai"
    assert result["metadata"]["ai_applied"] is True


def test_apply_ai_workflow_replaces_list(tmp_path, monkeypatch):
    import asyncio

    asyncio.run(_run_apply_ai_workflow(monkeypatch, tmp_path))
