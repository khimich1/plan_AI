"""FBS-003: parse FBS order lines (mark, grade, qty) + merge."""

from __future__ import annotations

from core.fbs_line_parser import (
    merge_fbs_lines,
    parse_fbs_line,
    parse_fbs_text,
    preserve_display_mark,
)


def test_parse_fbs_line_variants() -> None:
    cases = [
        ("ФБС 9.3.6-Т 2", "ФБС 9.3.6-Т", "B25", 2),
        ("ФБС 9.3.6-Т B25 2", "ФБС 9.3.6-Т", "B25", 2),
        ("ФБС 12.4.6-Т 25 1", "ФБС 12.4.6-Т", "B25", 1),
        ("ФБС 24.6.6-Т B7.5 3", "ФБС 24.6.6-Т", "B7_5", 3),
        ("фбс 9.5.6-т 5", "фбс 9.5.6-т", "B25", 5),
        ("ФБС 12.4.3-Т B20 1", "ФБС 12.4.3-Т", "B20", 1),
        ("ФБС 9.3.6-Т B22.5 4", "ФБС 9.3.6-Т", "B22_5", 4),
        ("ФБС9.3.6-Т", "ФБС 9.3.6-Т", "B25", 1),
    ]
    for line, mark, grade, qty in cases:
        result = parse_fbs_line(line)
        assert result.parsed is True, (line, result.reason_code)
        assert result.mark == mark
        assert result.concrete_grade == grade
        assert result.qty == qty


def test_preserve_display_mark() -> None:
    assert preserve_display_mark("ФБС 9.3.6-Т") == "ФБС 9.3.6-Т"
    assert preserve_display_mark("ФБС9.3.6-Т") == "ФБС 9.3.6-Т"
    assert preserve_display_mark("ФБС  9 . 3 . 6 - Т") == "ФБС 9.3.6-Т"


def test_parse_fbs_line_empty_and_invalid() -> None:
    empty = parse_fbs_line("")
    assert empty.parsed is False
    assert empty.reason_code == "empty_line"

    pile = parse_fbs_line("С120.35-12 2")
    assert pile.parsed is False

    bridge = parse_fbs_line("C8-35T1 2")
    assert bridge.parsed is False


def test_merge_same_lookup_mark_and_grade() -> None:
    lines = [
        parse_fbs_line("ФБС 9.3.6-Т 2"),
        parse_fbs_line("фбс 9.3.6-т B25 3"),
    ]
    merged = merge_fbs_lines(lines)
    assert len(merged) == 1
    assert merged[0].qty == 5
    assert merged[0].mark == "ФБС 9.3.6-Т"


def test_merge_different_grades_stay_separate() -> None:
    lines = [
        parse_fbs_line("ФБС 9.3.6-Т B25 2"),
        parse_fbs_line("ФБС 9.3.6-Т B7.5 3"),
    ]
    merged = merge_fbs_lines(lines)
    assert len(merged) == 2


def test_parse_fbs_text_multiline() -> None:
    text = "ФБС 9.3.6-Т 2\nФБС 12.4.6-Т B20 1\nФБС 24.6.6-Т B25 4"
    results = parse_fbs_text(text)
    assert len(results) == 3
    by_mark = {r.mark: r for r in results}
    assert by_mark["ФБС 9.3.6-Т"].qty == 2
    assert by_mark["ФБС 12.4.6-Т"].concrete_grade == "B20"
    assert by_mark["ФБС 24.6.6-Т"].qty == 4
