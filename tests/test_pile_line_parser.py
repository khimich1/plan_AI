"""PILE-002: parse pile order lines (mark, grade, qty) + merge duplicates."""

from __future__ import annotations

import pytest

from core.pile_line_parser import (
    PileLineParseResult,
    merge_pile_lines,
    parse_pile_line,
    parse_pile_text,
)

# Marks from tests/test_pile_price_import.py fixtures
MARK_12 = "С120.35-12"
MARK_13I = "С120.35-13и"


@pytest.mark.parametrize(
    ("line", "mark", "grade", "qty"),
    [
        (f"{MARK_12} 5", MARK_12, "B25", 5),
        (f"{MARK_12} B25 5", MARK_12, "B25", 5),
        (f"{MARK_12} 25 5", MARK_12, "B25", 5),
        (f"{MARK_13I} B30 2", MARK_13I, "B30_granite", 2),
        ("с 120.35-12 b25 5", MARK_12, "B25", 5),
        (f"{MARK_12}", MARK_12, "B25", 1),
        (f"{MARK_12} B20 3", MARK_12, "B20", 3),
        (f"{MARK_12} 22.5 4", MARK_12, "B22_5", 4),
        ("С110.30-9у 3", "С110.30-9у", "B25", 3),
        ("C 110.30-9у 2", "С110.30-9у", "B25", 2),
        ("С110.30-9.1у 1", "С110.30-9.1у", "B25", 1),
        ("С110.30-6у", "С110.30-6у", "B25", 1),
        ("С60.30.6 B22.5 6", "С60.30-6", "B22_5", 6),
        ("С30.30.6 14", "С30.30-6", "B25", 14),
        ("C60.30.6 B22.5 6", "С60.30-6", "B22_5", 6),
        ("С110.30.9.1у 1", "С110.30-9.1у", "B25", 1),
        ("С60.30.13-22 2", "С60.30-13-22", "B25", 2),
    ],
)
def test_parse_pile_line(line: str, mark: str, grade: str, qty: int) -> None:
    result = parse_pile_line(line)
    assert result.parsed is True
    assert result.mark == mark
    assert result.concrete_grade == grade
    assert result.qty == qty


def test_reinforced_flag_on_u_suffix() -> None:
    assert parse_pile_line("С110.30-9у").reinforced is True
    assert parse_pile_line("С110.30-9").reinforced is False
    assert parse_pile_line("С120.35-13и").reinforced is False


def test_parse_pile_line_empty_and_invalid() -> None:
    empty = parse_pile_line("")
    assert empty.parsed is False
    assert empty.reason_code == "empty_line"

    bad = parse_pile_line("ПБ 78-12-8п 2")
    assert bad.parsed is False


def test_merge_same_mark_and_grade_sums_qty() -> None:
    lines = [
        parse_pile_line(f"{MARK_12} 5"),
        parse_pile_line(f"{MARK_12} B25 3"),
    ]
    merged = merge_pile_lines(lines)
    assert len(merged) == 1
    assert merged[0].mark == MARK_12
    assert merged[0].concrete_grade == "B25"
    assert merged[0].qty == 8


def test_merge_different_grades_stay_separate() -> None:
    lines = [
        parse_pile_line(f"{MARK_12} B25 5"),
        parse_pile_line(f"{MARK_12} B20 3"),
    ]
    merged = merge_pile_lines(lines)
    assert len(merged) == 2
    by_grade = {m.concrete_grade: m.qty for m in merged}
    assert by_grade == {"B25": 5, "B20": 3}


def test_merge_plain_and_reinforced_stay_separate() -> None:
    lines = [
        parse_pile_line("С110.30-9 2"),
        parse_pile_line("С110.30-9у 3"),
    ]
    merged = merge_pile_lines(lines)
    assert len(merged) == 2
    by_mark = {item.mark: item for item in merged}
    assert by_mark["С110.30-9"].qty == 2
    assert by_mark["С110.30-9у"].qty == 3
    assert by_mark["С110.30-9"].reinforced is False
    assert by_mark["С110.30-9у"].reinforced is True


def test_parse_pile_text_multiline_with_merge() -> None:
    text = f"{MARK_12} 5\n{MARK_12} B25 3\n{MARK_13I} B30 2"
    results = parse_pile_text(text)
    assert len(results) == 2
    by_mark = {r.mark: r for r in results}
    assert by_mark[MARK_12].qty == 8
    assert by_mark[MARK_13I].qty == 2
    assert by_mark[MARK_13I].concrete_grade == "B30_granite"


def test_merge_pile_lines_skips_unparsed() -> None:
    merged = merge_pile_lines(
        [
            parse_pile_line(f"{MARK_12} 2"),
            PileLineParseResult(parsed=False, reason_code="bad"),
        ]
    )
    assert len(merged) == 1
    assert merged[0].qty == 2
