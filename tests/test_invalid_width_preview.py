from __future__ import annotations

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_draft_service import CommercialDraftService
from app.services.commercial_service import CommercialPreviewResult, CommercialService
from core.plate_order_context import PlateOrderContext


def _preview(*, invalid_width_lines: list[dict], wide_plate_lines: list | None = None) -> CommercialPreviewResult:
    order = PlateOrder()
    return CommercialPreviewResult(
        parse_result=ParseResult(
            order=order,
            normalized_text="",
            normalized_lines=[],
            wide_plate_lines=wide_plate_lines or [],
        ),
        optimization_context=OptimizationContext(order=order),
        order_data=[],
        price_rows=[],
        breakdown_tables=[],
        total_sum=0.0,
        invalid_width_lines=invalid_width_lines,
    )


def test_serialize_and_preview_metadata_sets_resolved_flag() -> None:
    draft_service = CommercialDraftService()
    lines = [
        {
            "id": "invalid-width-1",
            "name": "Плиты ПБ 29-8-8п",
            "line": "ПБ 29-8-8п 1",
            "qty": 1,
            "length_m": 2.9,
            "width_m": 0.8,
            "width_mm": 800,
            "load_class": 800,
            "replacements": [
                {"width_mm": 720, "width_label": "7,2", "price": 100.0},
                {"width_mm": 860, "width_label": "8,6"},
            ],
        }
    ]
    metadata = draft_service.build_preview_metadata(
        preview=_preview(invalid_width_lines=lines),
        base_metadata={},
        source_type="text",
        original_text="",
        ocr_text="",
        input_text="ПБ 29-8-8п 1",
        last_source_filename="",
        plate_batches=[],
        wide_plates_resolved=True,
        source_metadata={},
    )
    assert metadata["invalid_widths_resolved"] is False
    assert len(metadata["invalid_width_lines"]) == 1
    assert metadata["invalid_width_lines"][0]["replacements"][0]["width_mm"] == 720
    assert "price" not in metadata["invalid_width_lines"][0]["replacements"][1]


def test_preview_metadata_empty_lines_resolved_true() -> None:
    draft_service = CommercialDraftService()
    metadata = draft_service.build_preview_metadata(
        preview=_preview(invalid_width_lines=[]),
        base_metadata={},
        source_type="text",
        original_text="",
        ocr_text="",
        input_text="ПБ 29-12-8п 1",
        last_source_filename="",
        plate_batches=[],
        wide_plates_resolved=True,
        source_metadata={},
    )
    assert metadata["invalid_width_lines"] == []
    assert metadata["invalid_widths_resolved"] is True


def test_pile_preview_metadata_has_no_invalid_width_gate() -> None:
    class _PilePreview:
        warnings: list[str] = []
        unparsed_lines: list[str] = []
        normalized_text = "C80.30-11 1"
        normalized_lines = ["C80.30-11 1"]
        order_data: list[dict] = []
        total_sum = 0.0

    metadata = CommercialDraftService().build_pile_preview_metadata(
        preview=_PilePreview(),
        base_metadata={},
        source_type="text",
        original_text="",
        ocr_text="",
        input_text="C80.30-11 1",
        last_source_filename="",
        pile_batches=[],
        source_metadata={},
    )
    assert metadata["invalid_width_lines"] == []
    assert metadata["invalid_widths_resolved"] is True


def test_generate_preview_screen_mix_three_invalid_eights() -> None:
    service = CommercialService()
    preview = service.generate_preview(
        text=(
            "ПБ 29-12-8п 1\n"
            "ПБ 29-8-8п 1\n"
            "ПБ 32-8-8п 1\n"
            "ПБ 36-8-8п 1\n"
            "ПБ 36-12-8п 1\n"
        ),
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )
    marks = [item["name"] + " " + item["line"] for item in preview.invalid_width_lines]
    joined = " ".join(marks)
    assert len(preview.invalid_width_lines) == 3
    assert "29-8" in joined
    assert "32-8" in joined
    assert "36-8" in joined
    assert "29-12" not in joined
    assert "36-12" not in joined


def test_generate_preview_skips_zero_three_and_wide_fifteen() -> None:
    service = CommercialService()
    preview = service.generate_preview(
        text="ПБ 78-0.3-8п 1\nПБ 78-3-8п 1\nПБ 60-15-8п 1\nПБ 60-12-8п 1\n",
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )
    assert preview.invalid_width_lines == []
    assert preview.parse_result.wide_plate_lines
