import pytest

from app.services.plate_parser_service import PlateParserService
from core.dobor_split import expand_dobor_line
from core.plate_line_parser import parse_line
from core.plate_text_normalizer import normalize_order_text


@pytest.mark.parametrize(
    "marker",
    ["+ доб 5-шт", "+доб 5-шт", "доб 5-шт", "добор 5-шт", "+доб 5шт", "+доб 5 шт", "+доб 5"],
)
def test_expand_dobor_line_canonical_markers(marker: str):
    source = f"ПБ 57-7,2-8п {marker}"
    lines, pair, warnings = expand_dobor_line(source)

    assert warnings == []
    assert lines == ["ПБ 57-7,2-8п 5", "ПБ 57-4,8-8п 5"]
    assert pair is not None
    assert pair.pair_id == "dobor-1"
    assert pair.source_line == source
    assert pair.primary_line == "ПБ 57-7,2-8п 5"
    assert pair.complement_line == "ПБ 57-4,8-8п 5"


def test_expand_dobor_line_default_qty_one():
    lines, pair, warnings = expand_dobor_line("ПБ 57-7,2-8п +доб")

    assert warnings == []
    assert lines == ["ПБ 57-7,2-8п", "ПБ 57-4,8-8п"]
    assert pair is not None


def test_expand_dobor_line_without_marker_returns_single_line():
    lines, pair, warnings = expand_dobor_line("ПБ 57-7,2-8п 5")

    assert pair is None
    assert warnings == []
    assert lines == ["ПБ 57-7,2-8п 5"]


def test_expand_dobor_line_qty_conflict_uses_dobor_qty_with_warning():
    lines, pair, warnings = expand_dobor_line("ПБ 57-7,2-8п 3 + доб 5-шт")

    assert lines == ["ПБ 57-7,2-8п 5", "ПБ 57-4,8-8п 5"]
    assert pair is not None
    assert len(warnings) == 1
    assert "Конфликт количества" in warnings[0]


def test_expand_dobor_line_width_twelve_no_split():
    lines, pair, warnings = expand_dobor_line("ПБ 57-12-8п + доб 5-шт")

    assert pair is None
    assert len(lines) == 1
    assert "+ доб" in lines[0] or "доб" in lines[0].lower()
    assert len(warnings) == 1
    assert "невозможен" in warnings[0]


def test_normalize_order_text_splits_dobor_line():
    result = normalize_order_text("ПБ 57-7,2-8п + доб 5-шт")

    assert result.normalized_lines == ["ПБ 57-7,2-8п 5", "ПБ 57-4,8-8п 5"]
    assert len(result.dobor_pairs) == 1
    assert result.dobor_pairs[0].primary_line == "ПБ 57-7,2-8п 5"
    assert result.dobor_pairs[0].complement_line == "ПБ 57-4,8-8п 5"


def test_parse_plate_text_dobor_produces_two_positions():
    service = PlateParserService()
    result = service.parse_plate_text("ПБ 57-7,2-8п + доб 5-шт")

    assert not result.unparsed_lines
    assert result.normalized_lines == ["ПБ 57-7,2-8п 5", "ПБ 57-4,8-8п 5"]
    assert len(result.dobor_pairs) == 1

    primary = parse_line("ПБ 57-7,2-8п 5")
    complement = parse_line("ПБ 57-4,8-8п 5")

    assert primary.parsed and complement.parsed
    assert primary.length_m == pytest.approx(5.7)
    assert complement.length_m == pytest.approx(5.7)
    assert primary.width_m == pytest.approx(0.72)
    assert complement.width_m == pytest.approx(0.48)
    assert primary.qty == 5
    assert complement.qty == 5
    assert primary.load_code == 8.0
    assert complement.load_code == 8.0

    assert sum(result.order.plate_load_details.values()) == 10


def test_commercial_draft_metadata_serializes_dobor_pairs():
    from app.services.commercial_draft_service import CommercialDraftService
    from core.dobor_split import DoborPair

    pairs = [
        DoborPair(
            pair_id="dobor-1",
            primary_line="ПБ 57-7,2-8п 5",
            complement_line="ПБ 57-4,8-8п 5",
            source_line="ПБ 57-7,2-8п + доб 5-шт",
        )
    ]
    serialized = CommercialDraftService.serialize_dobor_pairs(pairs)

    assert serialized == [
        {
            "id": "dobor-1",
            "source_line": "ПБ 57-7,2-8п + доб 5-шт",
            "primary_line": "ПБ 57-7,2-8п 5",
            "complement_line": "ПБ 57-4,8-8п 5",
        }
    ]

