"""STEP-005: step OCR prompt + text normalizer."""

from __future__ import annotations

from core.step_format_prompt import build_step_parser_system_prompt
from core.step_text_normalizer import normalize_step_order_text


def test_build_step_parser_system_prompt_non_empty() -> None:
    prompt = build_step_parser_system_prompt()
    assert prompt.strip()
    assert "ЛС" in prompt
    assert "JSON" in prompt
    assert "бетон" in prompt.lower() or "B15" in prompt


def test_normalize_step_order_text_splits_multiline() -> None:
    text = "ЛС11 10\nЛестничные ступени ЛС14-1лев  5"
    result = normalize_step_order_text(text)
    assert len(result.normalized_lines) == 2
    assert result.normalized_lines[0].startswith("ЛС11")
    assert "ЛС14-1ЛЕВ" in result.normalized_lines[1].upper().replace("Ё", "Е")
    assert "Лестничные" not in result.normalized_lines[1]
    assert result.normalized_text.count("\n") == 1
