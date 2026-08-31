"""BP-005: bridge pile format prompt + text normalizer."""

from __future__ import annotations

from core.bridge_pile_format_prompt import build_bridge_pile_parser_system_prompt
from core.bridge_pile_text_normalizer import normalize_bridge_pile_order_text


def test_bridge_pile_format_prompt_non_empty() -> None:
    prompt = build_bridge_pile_parser_system_prompt()
    assert "мостов" in prompt.lower()
    assert "B25" in prompt
    assert "B30" in prompt
    assert "C8-35T1" in prompt
    assert "JSON" in prompt


def test_normalize_bridge_pile_order_text_basic() -> None:
    result = normalize_bridge_pile_order_text("C8-35T1  2\nC8-35В4 1 шт")
    assert result.normalized_lines == ["C8-35T1 2", "C8-35В4 1"]
    assert "C8-35В4" in result.normalized_text
