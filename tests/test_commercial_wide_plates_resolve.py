from __future__ import annotations

from typing import Any

import pytest

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_service import CommercialPreviewResult
from app.services.commercial_workflow_service import CommercialWorkflowService
from core.plate_order_context import PlateOrderContext


def test_resolve_wide_plates_exclude_compact_without_pe_matches_display_line(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """OCR compact «68-15-8 2» must exclude when metadata line is «ПБ 68-15-8п 2»."""
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "original_text": "68-15-8 2\n68-12-8 1",
            "ocr_text": "",
            "input_text": "68-15-8 2\n68-12-8 1",
            "normalized_lines": ["68-15-8 2", "68-12-8 1"],
            "plate_batches": [
                {
                    "source_type": "text",
                    "original_text": "",
                    "normalized_text": "68-15-8 2\n68-12-8 1",
                    "ocr_text": "",
                    "filename": "",
                }
            ],
            "wide_plate_lines": [{"id": "wide-1", "line": "ПБ 68-15-8п 2", "qty": 2}],
            "last_source_filename": "",
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)
    captured: dict[str, str] = {}

    def fake_generate_preview(
        *,
        text: str | None = None,
        parse_result: ParseResult | None = None,
        plate_order_ctx: Any = None,
    ) -> CommercialPreviewResult:
        preview_text = text or ""
        captured["text"] = preview_text
        return CommercialPreviewResult(
            parse_result=ParseResult(
                order=PlateOrder(),
                normalized_text=preview_text,
                normalized_lines=[line for line in preview_text.splitlines() if line.strip()],
            ),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=[],
            price_rows=[],
            breakdown_tables=[],
            total_sum=0.0,
        )

    monkeypatch.setattr(workflow.commercial_service, "generate_preview", fake_generate_preview)
    saved: dict[str, Any] = {}

    def fake_replace_preview(draft_id: str, **kwargs: Any) -> str:
        saved["metadata"] = kwargs.get("metadata")
        return draft_id

    monkeypatch.setattr(workflow.draft_store, "replace_preview", fake_replace_preview)
    monkeypatch.setattr(workflow, "_persist_wizard_step", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        workflow,
        "get_draft_details",
        lambda draft_id: {"draft_id": draft_id, "metadata": saved.get("metadata", {})},
    )

    workflow.resolve_wide_plates(
        "draft-1",
        decisions=[{"line_id": "wide-1", "action": "exclude"}],
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )

    assert captured["text"] == "68-12-8 1"
    assert "68-15-8" not in captured["text"]
    assert saved["metadata"]["plate_batches"][0]["normalized_text"] == "68-12-8 1"
