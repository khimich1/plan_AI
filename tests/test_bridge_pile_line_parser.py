"""BP-003: parse bridge-pile order lines (mark, grade, qty) + merge."""

from __future__ import annotations

import pytest

from core.bridge_pile_line_parser import (
    merge_bridge_pile_lines,
    parse_bridge_pile_line,
    parse_bridge_pile_text,
    preserve_display_mark,
)


@pytest.mark.parametrize(
    ("line", "mark", "grade", "qty"),
    [
        ("C8-35T1 2", "C8-35T1", "B25", 2),
        ("C8-35T1 B25 2", "C8-35T1", "B25", 2),
        ("C8-35В4 25 1", "C8-35В4", "B25", 1),
        ("C13-35T4 B30 3", "C13-35T4", "B30", 3),
        ("с8-35t1 5", "с8-35t1", "B25", 5),
        ("С7-35Т5", "С7-35Т5", "B25", 1),
        ("C10-35B7 B30 2", "C10-35B7", "B30", 2),
    ],
)
def test_parse_bridge_pile_line(line: str, mark: str, grade: str, qty: int) -> None:
    result = parse_bridge_pile_line(line)
    assert result.parsed is True
    assert result.mark == mark
    assert result.concrete_grade == grade
    assert result.qty == qty


def test_preserve_display_mark_keeps_cyrillic() -> None:
    assert preserve_display_mark("C8-35В4") == "C8-35В4"
    assert preserve_display_mark("С7-35Т5") == "С7-35Т5"
    assert preserve_display_mark("C 8-35T1") == "C8-35T1"


def test_parse_bridge_pile_line_empty_and_invalid() -> None:
    empty = parse_bridge_pile_line("")
    assert empty.parsed is False
    assert empty.reason_code == "empty_line"

    bad = parse_bridge_pile_line("С120.35-12 2")
    assert bad.parsed is False

    plate = parse_bridge_pile_line("ПБ 78-12-8п 2")
    assert plate.parsed is False


def test_merge_same_lookup_mark_and_grade() -> None:
    lines = [
        parse_bridge_pile_line("C8-35T1 2"),
        parse_bridge_pile_line("c8-35t1 B25 3"),
    ]
    merged = merge_bridge_pile_lines(lines)
    assert len(merged) == 1
    assert merged[0].qty == 5
    # Keep first display spelling
    assert merged[0].mark == "C8-35T1"


def test_merge_different_grades_stay_separate() -> None:
    lines = [
        parse_bridge_pile_line("C8-35T1 B25 2"),
        parse_bridge_pile_line("C8-35T1 B30 3"),
    ]
    merged = merge_bridge_pile_lines(lines)
    assert len(merged) == 2


def test_parse_bridge_pile_text_multiline() -> None:
    text = "C8-35T1 2\nC8-35В4 B25 1\nC13-40T3 B30 4"
    results = parse_bridge_pile_text(text)
    assert len(results) == 3
    by_mark = {r.mark: r for r in results}
    assert by_mark["C8-35T1"].qty == 2
    assert by_mark["C8-35В4"].qty == 1
    assert by_mark["C13-40T3"].concrete_grade == "B30"
