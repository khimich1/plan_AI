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


OCR_WIDE_LINE = "ПБ 76-15-8 13"
CANON_WIDE_LINE = "ПБ 76-15-8п 13"
WIDE_REPLACEMENT_TEXT = "ПБ 76-12-8 13\nПБ 76-3.0-8 13"
IMAGE_KEEP_LINE = "ПБ 60-12-8 1"


def _image_ocr_wide_metadata(*, keep_line: bool) -> dict[str, Any]:
    """Image batch stays raw OCR; wide_plate_lines keep the parser «п» mark."""
    lines = [OCR_WIDE_LINE, IMAGE_KEEP_LINE] if keep_line else [OCR_WIDE_LINE]
    text = "\n".join(lines)
    return {
        "source_type": "image",
        "original_text": text,
        "ocr_text": text,
        "input_text": text,
        "normalized_lines": list(lines),
        "plate_batches": [
            {
                "source_type": "image",
                "original_text": text,
                "normalized_text": text,
                "ocr_text": text,
                "filename": "photo.png",
            }
        ],
        "wide_plate_lines": [{"id": "wide-1", "line": CANON_WIDE_LINE, "qty": 13}],
        "last_source_filename": "photo.png",
    }


def _resolve_wide_plates_with_fake_preview(
    monkeypatch: pytest.MonkeyPatch,
    *,
    metadata: dict[str, Any],
    decisions: list[dict[str, Any]],
) -> dict[str, Any]:
    """Rewrite check only: preview and persistence are fakes, no live DB."""
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": metadata,
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)

    def fake_generate_preview(
        *,
        text: str | None = None,
        parse_result: ParseResult | None = None,
        plate_order_ctx: Any = None,
    ) -> CommercialPreviewResult:
        del parse_result, plate_order_ctx
        preview_text = text or ""
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
        decisions=decisions,
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )
    assert "metadata" in saved
    return saved


def test_resolve_wide_plates_image_replace_without_pe_rewrites_batch(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """OCR «ПБ 76-15-8 13» replaces when the wide line is «ПБ 76-15-8п 13»."""
    saved = _resolve_wide_plates_with_fake_preview(
        monkeypatch,
        metadata=_image_ocr_wide_metadata(keep_line=False),
        decisions=[
            {
                "line_id": "wide-1",
                "source_line": CANON_WIDE_LINE,
                "action": "replace",
                "replacement_text": WIDE_REPLACEMENT_TEXT,
            }
        ],
    )
    metadata = saved["metadata"]
    batch_text = metadata["plate_batches"][0]["normalized_text"]
    for text in (metadata["input_text"], batch_text):
        assert "ПБ 76-12-8 13" in text
        assert "ПБ 76-3.0-8 13" in text
        assert "76-15-8" not in text


def test_resolve_wide_plates_image_exclude_without_pe_removes_mark(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Exclude drops the wide OCR mark from input_text and the image batch."""
    saved = _resolve_wide_plates_with_fake_preview(
        monkeypatch,
        metadata=_image_ocr_wide_metadata(keep_line=True),
        decisions=[{"line_id": "wide-1", "source_line": CANON_WIDE_LINE, "action": "exclude"}],
    )
    metadata = saved["metadata"]
    batch_text = metadata["plate_batches"][0]["normalized_text"]
    assert metadata["input_text"] == IMAGE_KEEP_LINE
    assert batch_text == IMAGE_KEEP_LINE
    assert "76-15-8" not in metadata["input_text"]
    assert "76-15-8" not in batch_text


def test_align_text_batch_normalized_uses_parser_text() -> None:
    from app.services.product_draft_handler import ProductDraftHandler
    from types import SimpleNamespace

    preview = SimpleNamespace(parse_result=SimpleNamespace(normalized_text="ПБ 65-14-8п"))
    batches = [{"source_type": "text", "normalized_text": "ПБ 65-14-8"}]
    ProductDraftHandler._align_text_batch_normalized(batches, "text", preview)
    assert batches[0]["normalized_text"] == "ПБ 65-14-8п"

    image_batches = [{"source_type": "image", "normalized_text": "raw ocr"}]
    ProductDraftHandler._align_text_batch_normalized(image_batches, "image", preview)
    assert image_batches[0]["normalized_text"] == "raw ocr"
