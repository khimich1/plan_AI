"""PILE-004: pile OCR prompt + text normalizer."""

from __future__ import annotations

from core.pile_format_prompt import build_pile_parser_system_prompt
from core.pile_text_normalizer import normalize_pile_order_text


def test_build_pile_parser_system_prompt_non_empty() -> None:
    prompt = build_pile_parser_system_prompt()
    assert prompt.strip()
    assert "С120" in prompt or "сва" in prompt.lower()
    assert "JSON" in prompt


def test_normalize_pile_order_text_splits_multiline() -> None:
    text = "С120.35-12 B25 5\nс 120.35-13и  b30  2"
    result = normalize_pile_order_text(text)
    assert len(result.normalized_lines) == 2
    assert "С120.35-12" in result.normalized_lines[0]
    assert "B25" in result.normalized_lines[0].upper()
    assert result.normalized_text.count("\n") == 1
