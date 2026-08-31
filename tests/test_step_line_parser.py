"""STEP-003: parse stair-step order lines (mark + qty) + merge duplicates."""

from __future__ import annotations

import pytest

from core.step_line_parser import (
    StepLineParseResult,
    merge_step_lines,
    parse_step_line,
    parse_step_text,
)

# Marks covering suffix families from the ЛС price list
MARKS = [
    "ЛС11",
    "ЛС11-1",
    "ЛС11-1ЛЕВ",
    "ЛС11-2",
    "ЛС11-2ЛЕВ",
    "ЛС11-Б",
    "ЛС11-Б-1",
    "ЛС11-Б-1ЛЕВ",
    "ЛС14-1ЛЕВ",
    "ЛС14-Б",
    "ЛС22-2ЛЕВ",
]


@pytest.mark.parametrize(
    ("line", "mark", "qty"),
    [
        ("ЛС11 10", "ЛС11", 10),
        ("ЛС14-1лев 5", "ЛС14-1ЛЕВ", 5),
        ("ЛС11-Б-1 2", "ЛС11-Б-1", 2),
        ("лс12-2лев 3", "ЛС12-2ЛЕВ", 3),
        ("Лестничные ступени ЛС15-1 4", "ЛС15-1", 4),
        ("ЛС11", "ЛС11", 1),
        ("ЛС22-2лев 8 шт", "ЛС22-2ЛЕВ", 8),
        ("Лестничные ступени  ЛС14-Б 2", "ЛС14-Б", 2),
    ],
)
def test_parse_step_line(line: str, mark: str, qty: int) -> None:
    result = parse_step_line(line)
    assert result.parsed is True
    assert result.mark == mark
    assert result.qty == qty


@pytest.mark.parametrize("mark", MARKS)
def test_parse_all_price_list_suffix_families(mark: str) -> None:
    result = parse_step_line(f"{mark} 2")
    assert result.parsed is True
    assert result.mark == mark
    assert result.qty == 2


def test_parse_step_line_empty_and_invalid() -> None:
    empty = parse_step_line("")
    assert empty.parsed is False
    assert empty.reason_code == "empty_line"

    bad = parse_step_line("ПБ 78-12-8п 2")
    assert bad.parsed is False

    pile = parse_step_line("С120.35-12 5")
    assert pile.parsed is False


def test_merge_same_mark_sums_qty() -> None:
    lines = [
        parse_step_line("ЛС11 5"),
        parse_step_line("ЛС11 3"),
    ]
    merged = merge_step_lines(lines)
    assert len(merged) == 1
    assert merged[0].mark == "ЛС11"
    assert merged[0].qty == 8


def test_merge_different_marks_stay_separate() -> None:
    lines = [
        parse_step_line("ЛС11 5"),
        parse_step_line("ЛС14-1лев 3"),
    ]
    merged = merge_step_lines(lines)
    assert len(merged) == 2
    by_mark = {m.mark: m.qty for m in merged}
    assert by_mark == {"ЛС11": 5, "ЛС14-1ЛЕВ": 3}


def test_parse_step_text_multiline_with_merge() -> None:
    text = "ЛС11 5\nЛС11 3\nЛС14-1лев 2"
    results = parse_step_text(text)
    assert len(results) == 2
    by_mark = {r.mark: r for r in results}
    assert by_mark["ЛС11"].qty == 8
    assert by_mark["ЛС14-1ЛЕВ"].qty == 2


def test_merge_step_lines_skips_unparsed() -> None:
    merged = merge_step_lines(
        [
            parse_step_line("ЛС11 2"),
            StepLineParseResult(parsed=False, reason_code="bad"),
        ]
    )
    assert len(merged) == 1
    assert merged[0].qty == 2
