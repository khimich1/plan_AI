#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse FBS order lines: mark, concrete grade, quantity.

Display mark = manager spelling (trim/spaces only). Lookup uses separate normalizer.
Grades: B7_5 / B20 / B22_5 / B25.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace

from core.fbs_price_db import (
    FBS_GRADE_CODES,
    grade_code_from_fbs_value,
    normalize_fbs_mark_for_lookup,
)

DEFAULT_CONCRETE_GRADE = "B25"

# ФБС 9.3.6-Т / ФБС9.3.6-T / фбс 24.6.6-т
_FBS_MARK_RE = re.compile(
    r"^(ФБС\s*\d+\s*\.\s*\d+\s*\.\s*\d+\s*-\s*[TТ])",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class FbsLineParseResult:
    parsed: bool
    mark: str = ""
    concrete_grade: str | None = None
    qty: int = 1
    reason_code: str = ""
    reason_text: str = ""


def preserve_display_mark(raw_mark: str) -> str:
    """Keep manager spelling; only collapse whitespace around dots/dash."""
    mark = (raw_mark or "").strip()
    mark = re.sub(r"\s{2,}", " ", mark)
    # ФБС9.3.6-Т → ФБС 9.3.6-Т (space after prefix if missing)
    mark = re.sub(r"^(ФБС)\s*", r"\1 ", mark, flags=re.IGNORECASE)
    mark = re.sub(r"\s*\.\s*", ".", mark)
    mark = re.sub(r"\s*-\s*", "-", mark)
    return mark.strip()


def _parse_grade_and_qty(
    tokens: list[str],
    *,
    default_grade: str,
) -> tuple[str, int] | None:
    if not tokens:
        return default_grade, 1

    if len(tokens) == 1:
        token = tokens[0]
        grade = grade_code_from_fbs_value(token)
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
    grade = grade_code_from_fbs_value(grade_text) or grade_code_from_fbs_value(
        grade_tokens[0]
    )
    if grade is None:
        return None
    return grade, qty


def parse_fbs_line(
    raw_line: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> FbsLineParseResult:
    """Parse one FBS order line into mark, grade, and quantity."""
    line = (raw_line or "").strip()
    if not line:
        return FbsLineParseResult(
            parsed=False,
            reason_code="empty_line",
            reason_text="пустая строка",
        )

    match = _FBS_MARK_RE.match(line)
    if not match:
        return FbsLineParseResult(
            parsed=False,
            reason_code="pattern_not_matched",
            reason_text="не совпал формат строки ФБС",
        )

    mark = preserve_display_mark(match.group(1))
    remainder = line[match.end() :].strip()
    tokens = remainder.split() if remainder else []

    parsed_grade_qty = _parse_grade_and_qty(tokens, default_grade=default_grade)
    if parsed_grade_qty is None:
        return FbsLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="grade_qty_parse_failed",
            reason_text="не удалось распознать класс бетона или количество",
        )

    grade, qty = parsed_grade_qty
    if grade not in FBS_GRADE_CODES:
        return FbsLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="unknown_grade",
            reason_text=f"неизвестный класс бетона: {grade}",
        )

    return FbsLineParseResult(
        parsed=True,
        mark=mark,
        concrete_grade=grade,
        qty=qty,
    )


def merge_fbs_lines(
    lines: list[FbsLineParseResult],
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[FbsLineParseResult]:
    """Merge lines with the same lookup-mark + grade; keep first display spelling."""
    merged: dict[tuple[str, str], FbsLineParseResult] = {}
    for line in lines:
        if not line.parsed:
            continue
        grade = line.concrete_grade or default_grade
        key = (normalize_fbs_mark_for_lookup(line.mark), grade)
        if key in merged:
            existing = merged[key]
            merged[key] = replace(existing, qty=existing.qty + line.qty)
        else:
            merged[key] = replace(line, concrete_grade=grade)
    return list(merged.values())


def parse_fbs_text(
    text: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[FbsLineParseResult]:
    """Parse multiline FBS text and merge duplicate mark+grade rows."""
    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text or "") if part.strip()]
    parsed = [parse_fbs_line(line, default_grade=default_grade) for line in raw_lines]
    return merge_fbs_lines(parsed, default_grade=default_grade)
