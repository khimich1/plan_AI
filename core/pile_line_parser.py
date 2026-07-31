#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse pile order lines: mark, concrete grade, quantity."""

from __future__ import annotations

import re
from dataclasses import dataclass, replace

from core.pile_price_db import GRADE_CODES, grade_code_from_value

DEFAULT_CONCRETE_GRADE = "B25"

_PILE_MARK_RE = re.compile(
    r"^([СC]\s*[\d.,]+(?:-[\d.,]+(?:[иИ])?)?)",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class PileLineParseResult:
    parsed: bool
    mark: str = ""
    concrete_grade: str | None = None
    qty: int = 1
    reason_code: str = ""
    reason_text: str = ""


def _normalize_mark(raw_mark: str) -> str:
    mark = (raw_mark or "").strip()
    mark = re.sub(r"^[СC]\s*", "С", mark, flags=re.IGNORECASE)
    return mark


def _parse_grade_and_qty(
    tokens: list[str],
    *,
    default_grade: str,
) -> tuple[str, int] | None:
    if not tokens:
        return default_grade, 1

    if len(tokens) == 1:
        token = tokens[0]
        grade = grade_code_from_value(token)
        if grade:
            return grade, 1
        if token.isdigit():
            return default_grade, max(1, int(token))
        return None

    last = tokens[-1]
    if not last.isdigit():
        return None

    qty = max(1, int(last))
    grade_tokens = tokens[:-1]
    grade_text = " ".join(grade_tokens)
    grade = grade_code_from_value(grade_text) or grade_code_from_value(grade_tokens[0])
    if grade is None:
        return None
    return grade, qty


def parse_pile_line(
    raw_line: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> PileLineParseResult:
    """Parse one pile order line into mark, grade, and quantity."""
    line = (raw_line or "").strip()
    if not line:
        return PileLineParseResult(
            parsed=False,
            reason_code="empty_line",
            reason_text="пустая строка",
        )

    match = _PILE_MARK_RE.match(line)
    if not match:
        return PileLineParseResult(
            parsed=False,
            reason_code="pattern_not_matched",
            reason_text="не совпал формат строки сваи",
        )

    mark = _normalize_mark(match.group(1))
    remainder = line[match.end() :].strip()
    tokens = remainder.split() if remainder else []

    parsed_grade_qty = _parse_grade_and_qty(tokens, default_grade=default_grade)
    if parsed_grade_qty is None:
        return PileLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="grade_qty_parse_failed",
            reason_text="не удалось распознать класс бетона или количество",
        )

    grade, qty = parsed_grade_qty
    if grade not in GRADE_CODES:
        return PileLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="unknown_grade",
            reason_text=f"неизвестный класс бетона: {grade}",
        )

    return PileLineParseResult(
        parsed=True,
        mark=mark,
        concrete_grade=grade,
        qty=qty,
    )


def merge_pile_lines(
    lines: list[PileLineParseResult],
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[PileLineParseResult]:
    """Merge lines with the same mark+grade (Q16): sum quantities."""
    merged: dict[tuple[str, str], PileLineParseResult] = {}
    for line in lines:
        if not line.parsed:
            continue
        grade = line.concrete_grade or default_grade
        key = (line.mark, grade)
        if key in merged:
            existing = merged[key]
            merged[key] = replace(existing, qty=existing.qty + line.qty)
        else:
            merged[key] = replace(line, concrete_grade=grade)
    return list(merged.values())


def parse_pile_text(
    text: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[PileLineParseResult]:
    """Parse multiline pile text and merge duplicate mark+grade rows."""
    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text or "") if part.strip()]
    parsed = [parse_pile_line(line, default_grade=default_grade) for line in raw_lines]
    return merge_pile_lines(parsed, default_grade=default_grade)
