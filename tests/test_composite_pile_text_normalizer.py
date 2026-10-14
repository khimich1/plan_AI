"""CP-001/002: composite pile mark / order-text normalizer."""

from __future__ import annotations

import pytest

from core.composite_pile_text_normalizer import (
    canonicalize_composite_pile_mark,
    normalize_composite_pile_order_text,
)


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("Сваи С 60.30-ВС.1", "С60.30-ВС.1"),
        ("С 60.30-ВС.1", "С60.30-ВС.1"),
        ("С60.30-ВС.1", "С60.30-ВС.1"),
        ("C60.30-BC.1", "С60.30-ВС.1"),
        ("c 60.30-bc.1", "С60.30-ВС.1"),
        ("С60.30-Всв.1", "С60.30-ВСв.1"),
        ("С60.30-Нсв.1", "С60.30-НСв.1"),
        ("С60.30-ВСв4", "С60.30-ВСв.4"),
        ("С60.30-НСв6", "С60.30-НСв.6"),
        ("С140.30-С", "С140.30-С"),
        ("С140.30-Св", "С140.30-Св"),
        ("С140.30-Св.ВП", "С140.30-Св.ВП"),
        ("Сваи С140.30-Св.ВП", "С140.30-Св.ВП"),
    ],
)
def test_canonicalize_composite_pile_mark(raw: str, expected: str) -> None:
    assert canonicalize_composite_pile_mark(raw) == expected


def test_normalize_order_text_preserves_latin_grade() -> None:
    result = normalize_composite_pile_order_text("Сваи С 60.30-ВС.1 B25 4")
    assert result.normalized_lines == ["С60.30-ВС.1 B25 4"]
    assert "B25" in result.normalized_text
    assert "В25" not in result.normalized_text


def test_normalize_order_text_preserves_decimal_grade_token() -> None:
    result = normalize_composite_pile_order_text("С60.30-ВС.1 22.5 2")
    assert result.normalized_lines == ["С60.30-ВС.1 22.5 2"]


def test_normalize_order_text_multiline() -> None:
    text = "Сваи С 60.30-ВС.1 4\nC140.30-C 5"
    result = normalize_composite_pile_order_text(text)
    assert result.normalized_lines == ["С60.30-ВС.1 4", "С140.30-С 5"]
