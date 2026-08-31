#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse stair-march (ЛМ) order lines: mark, concrete grade, quantity."""

from __future__ import annotations

import re
from dataclasses import dataclass, replace

from core.march_price_db import GRADE_CODES, grade_code_from_value, normalize_march_mark
from core.pile_line_parser import DEFAULT_CONCRETE_GRADE

# 1ЛМ 27-11-14-4 [закладные справа] | ЛМ 2,8 | ЛМ 2.8
_MARCH_MARK_RE = re.compile(
    r"^("
    r"1ЛМ\s*\d+(?:\s*-\s*\d+)+|"  # 1ЛМ 27-11-14-4
    r"ЛМ\s*\d+[.,]\d+"  # ЛМ 2,8 / ЛМ 2.8
    r")"
    r"(?:\s+(закладные\s+справа))?",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class MarchLineParseResult:
    parsed: bool
    mark: str = ""
    concrete_grade: str | None = None
    qty: int = 1
    reason_code: str = ""
    reason_text: str = ""


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


def parse_march_line(
    raw_line: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> MarchLineParseResult:
    """Parse one march order line into mark, grade, and quantity."""
    line = (raw_line or "").strip()
    if not line:
        return MarchLineParseResult(
            parsed=False,
            reason_code="empty_line",
            reason_text="пустая строка",
        )

    # Strip optional full-name prefix before mark match
    line_for_match = re.sub(
        r"^лестничн(?:ые|ая|ый)?\s+марш[иае]?\s+",
        "",
        line,
        flags=re.IGNORECASE | re.UNICODE,
    ).strip()

    match = _MARCH_MARK_RE.match(line_for_match)
    if not match:
        return MarchLineParseResult(
            parsed=False,
            reason_code="pattern_not_matched",
            reason_text="не совпал формат строки марша",
        )

    raw_mark = match.group(1)
    if match.group(2):
        raw_mark = f"{raw_mark} {match.group(2)}"
    mark = normalize_march_mark(raw_mark)
    remainder = line_for_match[match.end() :].strip()
    tokens = remainder.split() if remainder else []

    parsed_grade_qty = _parse_grade_and_qty(tokens, default_grade=default_grade)
    if parsed_grade_qty is None:
        return MarchLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="grade_qty_parse_failed",
            reason_text="не удалось распознать класс бетона или количество",
        )

    grade, qty = parsed_grade_qty
    if grade not in GRADE_CODES:
        return MarchLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="unknown_grade",
            reason_text=f"неизвестный класс бетона: {grade}",
        )

    return MarchLineParseResult(
        parsed=True,
        mark=mark,
        concrete_grade=grade,
        qty=qty,
    )


def merge_march_lines(
    lines: list[MarchLineParseResult],
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[MarchLineParseResult]:
    """Merge lines with the same mark+grade: sum quantities."""
    merged: dict[tuple[str, str], MarchLineParseResult] = {}
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


def parse_march_text(
    text: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[MarchLineParseResult]:
    """Parse multiline march text and merge duplicate mark+grade rows."""
    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text or "") if part.strip()]
    parsed = [parse_march_line(line, default_grade=default_grade) for line in raw_lines]
    return merge_march_lines(parsed, default_grade=default_grade)
