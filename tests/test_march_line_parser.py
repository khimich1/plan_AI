"""MARCH-003: parse march order lines (mark, grade, qty) + merge duplicates."""

from __future__ import annotations

import pytest

from core.march_line_parser import (
    MarchLineParseResult,
    merge_march_lines,
    parse_march_line,
    parse_march_text,
)

MARK_1 = "1ЛМ 27-11-14-4"
MARK_EMBED = "1ЛМ 30-11-15-4 закладные справа"
MARK_LM28 = "ЛМ 2,8"


@pytest.mark.parametrize(
    ("line", "mark", "grade", "qty"),
    [
        (f"{MARK_1} 2", MARK_1, "B25", 2),
        (f"{MARK_1} B25 2", MARK_1, "B25", 2),
        (f"{MARK_1} 25 2", MARK_1, "B25", 2),
        (f"{MARK_EMBED} B22.5 1", MARK_EMBED, "B22_5", 1),
        ("ЛМ 2,8 5", MARK_LM28, "B25", 5),
        ("ЛМ 2.8 5", MARK_LM28, "B25", 5),
        (f"{MARK_1}", MARK_1, "B25", 1),
        (f"{MARK_1} B20 3", MARK_1, "B20", 3),
        ("Лестничные марши 1ЛМ 27-11-14-4 2", MARK_1, "B25", 2),
    ],
)
def test_parse_march_line(line: str, mark: str, grade: str, qty: int) -> None:
    result = parse_march_line(line)
    assert result.parsed is True
    assert result.mark == mark
    assert result.concrete_grade == grade
    assert result.qty == qty


def test_parse_march_line_empty_and_invalid() -> None:
    empty = parse_march_line("")
    assert empty.parsed is False
    assert empty.reason_code == "empty_line"

    bad = parse_march_line("ПБ 78-12-8п 2")
    assert bad.parsed is False


def test_merge_same_mark_and_grade_sums_qty() -> None:
    lines = [
        parse_march_line(f"{MARK_1} 2"),
        parse_march_line(f"{MARK_1} B25 3"),
    ]
    merged = merge_march_lines(lines)
    assert len(merged) == 1
    assert merged[0].mark == MARK_1
    assert merged[0].concrete_grade == "B25"
    assert merged[0].qty == 5


def test_merge_different_grades_stay_separate() -> None:
    lines = [
        parse_march_line(f"{MARK_1} B25 2"),
        parse_march_line(f"{MARK_1} B20 3"),
    ]
    merged = merge_march_lines(lines)
    assert len(merged) == 2
    by_grade = {m.concrete_grade: m.qty for m in merged}
    assert by_grade == {"B25": 2, "B20": 3}


def test_parse_march_text_multiline_with_merge() -> None:
    text = f"{MARK_1} 2\n{MARK_1} B25 3\nЛМ 2.8 5"
    results = parse_march_text(text)
    assert len(results) == 2
    by_mark = {r.mark: r for r in results}
    assert by_mark[MARK_1].qty == 5
    assert by_mark[MARK_LM28].qty == 5
    assert by_mark[MARK_LM28].concrete_grade == "B25"


def test_merge_march_lines_skips_unparsed() -> None:
    merged = merge_march_lines(
        [
            parse_march_line(f"{MARK_1} 2"),
            MarchLineParseResult(parsed=False, reason_code="bad"),
        ]
    )
    assert len(merged) == 1
    assert merged[0].qty == 2
