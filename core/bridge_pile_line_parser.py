#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse bridge-pile order lines: mark, concrete grade, quantity.

Display mark = manager spelling (trim/spaces only). Lookup uses separate normalizer.
Grades: B25 / B30 only.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace

from core.bridge_pile_price_db import (
    BRIDGE_PILE_GRADE_CODES,
    grade_code_from_bridge_value,
    normalize_bridge_pile_mark_for_lookup,
)

DEFAULT_CONCRETE_GRADE = "B25"

# C8-35T1 / С7-35Т5 / C8-35В4 / C10-35B7
_BRIDGE_PILE_MARK_RE = re.compile(
    r"^([СC]\s*\d+\s*-\s*\d+\s*[TТBВ]\s*\d+)",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class BridgePileLineParseResult:
    parsed: bool
    mark: str = ""
    concrete_grade: str | None = None
    qty: int = 1
    reason_code: str = ""
    reason_text: str = ""


def preserve_display_mark(raw_mark: str) -> str:
    """Keep manager spelling; only collapse whitespace."""
    mark = (raw_mark or "").strip()
    mark = re.sub(r"\s{2,}", " ", mark)
    # Drop internal spaces inside mark token (C 8-35T1 → C8-35T1) without latin/cyr rewrite
    mark = re.sub(r"\s*-\s*", "-", mark)
    mark = re.sub(r"([СC])\s+(\d)", r"\1\2", mark, flags=re.IGNORECASE)
    mark = re.sub(r"(\d)\s+([TТBВ])", r"\1\2", mark, flags=re.IGNORECASE)
    mark = re.sub(r"([TТBВ])\s+(\d)", r"\1\2", mark, flags=re.IGNORECASE)
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
        grade = grade_code_from_bridge_value(token)
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
    grade = grade_code_from_bridge_value(grade_text) or grade_code_from_bridge_value(
        grade_tokens[0]
    )
    if grade is None:
        return None
    return grade, qty


def parse_bridge_pile_line(
    raw_line: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> BridgePileLineParseResult:
    """Parse one bridge-pile order line into mark, grade, and quantity."""
    line = (raw_line or "").strip()
    if not line:
        return BridgePileLineParseResult(
            parsed=False,
            reason_code="empty_line",
            reason_text="пустая строка",
        )

    match = _BRIDGE_PILE_MARK_RE.match(line)
    if not match:
        return BridgePileLineParseResult(
            parsed=False,
            reason_code="pattern_not_matched",
            reason_text="не совпал формат строки мостовой сваи",
        )

    mark = preserve_display_mark(match.group(1))
    remainder = line[match.end() :].strip()
    tokens = remainder.split() if remainder else []

    parsed_grade_qty = _parse_grade_and_qty(tokens, default_grade=default_grade)
    if parsed_grade_qty is None:
        return BridgePileLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="grade_qty_parse_failed",
            reason_text="не удалось распознать класс бетона или количество",
        )

    grade, qty = parsed_grade_qty
    if grade not in BRIDGE_PILE_GRADE_CODES:
        return BridgePileLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="unknown_grade",
            reason_text=f"неизвестный класс бетона: {grade}",
        )

    return BridgePileLineParseResult(
        parsed=True,
        mark=mark,
        concrete_grade=grade,
        qty=qty,
    )


def merge_bridge_pile_lines(
    lines: list[BridgePileLineParseResult],
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[BridgePileLineParseResult]:
    """Merge lines with the same lookup-mark + grade; keep first display spelling."""
    merged: dict[tuple[str, str], BridgePileLineParseResult] = {}
    for line in lines:
        if not line.parsed:
            continue
        grade = line.concrete_grade or default_grade
        key = (normalize_bridge_pile_mark_for_lookup(line.mark), grade)
        if key in merged:
            existing = merged[key]
            merged[key] = replace(existing, qty=existing.qty + line.qty)
        else:
            merged[key] = replace(line, concrete_grade=grade)
    return list(merged.values())


def parse_bridge_pile_text(
    text: str,
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[BridgePileLineParseResult]:
    """Parse multiline bridge-pile text and merge duplicate mark+grade rows."""
    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text or "") if part.strip()]
    parsed = [
        parse_bridge_pile_line(line, default_grade=default_grade) for line in raw_lines
    ]
    return merge_bridge_pile_lines(parsed, default_grade=default_grade)
