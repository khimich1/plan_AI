from __future__ import annotations

from typing import Any

import pytest

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_service import CommercialPreviewResult
from app.services.commercial_workflow_service import CommercialWorkflowService
from core.plate_order_context import PlateOrderContext


def _fake_preview(text: str, order_data: list[dict[str, Any]]) -> CommercialPreviewResult:
    order = PlateOrder()
    lines = [line for line in text.split("\n") if line.strip()]
    return CommercialPreviewResult(
        parse_result=ParseResult(order=order, normalized_text=text, normalized_lines=lines),
        optimization_context=OptimizationContext(order=order),
        order_data=order_data,
        price_rows=[],
        breakdown_tables=[],
        total_sum=0.0,
        invalid_width_lines=[],
    )


def _invalid_item() -> dict[str, Any]:
    return {
        "id": "invalid-width-1",
        "name": "Плиты ПБ 29-8-8п",
        "line": "ПБ 29-8-8п 1",
        "qty": 1,
        "length_m": 2.9,
        "width_m": 0.8,
        "width_mm": 800,
        "load_class": 800,
        "replacements": [
            {"width_mm": 720, "width_label": "7,2", "price": 9000.0},
            {"width_mm": 860, "width_label": "8,6", "price": 10400.0},
        ],
    }


def test_resolve_invalid_widths_replace_to_86(monkeypatch: pytest.MonkeyPatch) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "original_text": "ПБ 29-8-8п 1",
            "ocr_text": "",
            "input_text": "ПБ 29-8-8п 1\nПБ 29-12-8п 1",
            "normalized_lines": ["ПБ 29-8-8п 1", "ПБ 29-12-8п 1"],
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "invalid_width_lines": [_invalid_item()],
            "invalid_widths_resolved": False,
            "plate_batches": [
                {
                    "source_type": "text",
                    "original_text": "",
                    "normalized_text": "ПБ 29-8-8п 1\nПБ 29-12-8п 1",
                    "ocr_text": "",
                    "filename": "",
                }
            ],
            "last_source_filename": "",
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)

    captured: dict[str, Any] = {}

    def fake_generate_preview(
        *,
        text: str | None = None,
        parse_result: ParseResult | None = None,
        plate_order_ctx: Any = None,
    ) -> CommercialPreviewResult:
        captured["text"] = text or ""
        return _fake_preview(
            text or "",
            [
                {
                    "name": "Плиты ПБ 29-8,6-8п",
                    "qty": 1,
                    "length_m": 2.9,
                    "width_m": 0.86,
                    "load_class": 800,
                    "unit_price": 10400.0,
                },
                {
                    "name": "Плиты ПБ 29-12-8п",
                    "qty": 1,
                    "length_m": 2.9,
                    "width_m": 1.2,
                    "load_class": 800,
                    "unit_price": 9000.0,
                },
            ],
        )

    monkeypatch.setattr(workflow.commercial_service, "generate_preview", fake_generate_preview)
    saved: dict[str, Any] = {}

    def fake_replace_preview(draft_id: str, **kwargs: Any) -> str:
        saved["metadata"] = kwargs.get("metadata")
        saved["order_data"] = kwargs.get("order_data")
        return draft_id

    monkeypatch.setattr(workflow.draft_store, "replace_preview", fake_replace_preview)
    monkeypatch.setattr(
        workflow,
        "get_draft_details",
        lambda draft_id: {
            "draft_id": draft_id,
            "metadata": saved["metadata"],
            "order_data": saved["order_data"],
        },
    )
    monkeypatch.setattr(workflow, "_persist_wizard_step", lambda *_args, **_kwargs: None)

    result = workflow.resolve_invalid_widths(
        "draft-1",
        decisions=[{"line_id": "invalid-width-1", "action": "replace_width", "width_mm": 860}],
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )

    assert captured["text"] == "ПБ 29-8,6-8п 1\nПБ 29-12-8п 1"
    assert result["metadata"]["invalid_widths_resolved"] is True
    assert result["metadata"]["invalid_width_decisions"]
    assert result["metadata"]["plate_batches"][0]["normalized_text"] == "ПБ 29-8,6-8п 1\nПБ 29-12-8п 1"


def test_resolve_invalid_widths_rejects_foreign_width(monkeypatch: pytest.MonkeyPatch) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "input_text": "ПБ 29-8-8п 1",
            "normalized_lines": ["ПБ 29-8-8п 1"],
            "wide_plates_resolved": True,
            "invalid_width_lines": [_invalid_item()],
            "invalid_widths_resolved": False,
            "plate_batches": [],
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)

    with pytest.raises(ValueError, match="не входит в предложенные"):
        workflow.resolve_invalid_widths(
            "draft-1",
            decisions=[{"line_id": "invalid-width-1", "action": "replace_width", "width_mm": 800}],
            plate_order_ctx=PlateOrderContext.fresh_empty(),
        )


def test_resolve_invalid_widths_exclude_last_only_line_fails(monkeypatch: pytest.MonkeyPatch) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "input_text": "ПБ 29-8-8п 1",
            "normalized_lines": ["ПБ 29-8-8п 1"],
            "wide_plates_resolved": True,
            "invalid_width_lines": [_invalid_item()],
            "invalid_widths_resolved": False,
            "plate_batches": [],
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)

    with pytest.raises(ValueError, match="список стал пустым"):
        workflow.resolve_invalid_widths(
            "draft-1",
            decisions=[{"line_id": "invalid-width-1", "action": "exclude"}],
            plate_order_ctx=PlateOrderContext.fresh_empty(),
        )


def test_resolve_invalid_widths_exclude_keeps_valid_twelves(monkeypatch: pytest.MonkeyPatch) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "input_text": "ПБ 29-8-8п 1\nПБ 29-12-8п 1",
            "normalized_lines": ["ПБ 29-8-8п 1", "ПБ 29-12-8п 1"],
            "wide_plates_resolved": True,
            "invalid_width_lines": [_invalid_item()],
            "invalid_widths_resolved": False,
            "plate_batches": [],
            "last_source_filename": "",
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)
    captured: dict[str, Any] = {}

    def fake_generate_preview(
        *,
        text: str | None = None,
        parse_result: ParseResult | None = None,
        plate_order_ctx: Any = None,
    ) -> CommercialPreviewResult:
        captured["text"] = text or ""
        return _fake_preview(
            text or "",
            [
                {
                    "name": "Плиты ПБ 29-12-8п",
                    "qty": 1,
                    "length_m": 2.9,
                    "width_m": 1.2,
                    "unit_price": 9000.0,
                }
            ],
        )

    monkeypatch.setattr(workflow.commercial_service, "generate_preview", fake_generate_preview)
    saved: dict[str, Any] = {}
    monkeypatch.setattr(
        workflow.draft_store,
        "replace_preview",
        lambda draft_id, **kwargs: saved.update(kwargs) or draft_id,
    )
    monkeypatch.setattr(
        workflow,
        "get_draft_details",
        lambda draft_id: {"draft_id": draft_id, "metadata": saved["metadata"], "order_data": saved["order_data"]},
    )
    monkeypatch.setattr(workflow, "_persist_wizard_step", lambda *_args, **_kwargs: None)

    result = workflow.resolve_invalid_widths(
        "draft-1",
        decisions=[{"line_id": "invalid-width-1", "action": "exclude"}],
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )
    assert captured["text"] == "ПБ 29-12-8п 1"
    assert result["metadata"]["invalid_widths_resolved"] is True
