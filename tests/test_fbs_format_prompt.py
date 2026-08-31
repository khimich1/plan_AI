"""FBS-005: FBS format prompt non-empty."""

from __future__ import annotations

from core.fbs_format_prompt import build_fbs_parser_system_prompt


def test_fbs_prompt_mentions_marks_and_grades() -> None:
    prompt = build_fbs_parser_system_prompt()
    assert "ФБС" in prompt
    assert "B25" in prompt
    assert "B7_5" in prompt or "B7.5" in prompt
    assert "JSON" in prompt
    assert len(prompt) > 200
