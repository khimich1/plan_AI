#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse stair-step (ЛС) order lines: mark + quantity (no concrete grade)."""

from __future__ import annotations

import re
from dataclasses import dataclass, replace

from core.step_price_db import extract_step_mark, normalize_step_mark


@dataclass
class StepLineParseResult:
    parsed: bool
    mark: str = ""
    qty: int = 1
    reason_code: str = ""
    reason_text: str = ""


def _parse_qty(tokens: list[str]) -> int | None:
    """Qty is an explicit trailing integer; bare mark → qty=1. Ignore B15/15 as grade."""
    if not tokens:
        return 1
    if len(tokens) == 1:
        token = tokens[0]
        # Explicit qty only when clearly an integer count (not confused with grade alone
        # when there is no other context — for steps we treat a single trailing int as qty).
        if token.isdigit():
            return max(1, int(token))
        return None
    last = tokens[-1]
    if not last.isdigit():
        return None
    return max(1, int(last))


def parse_step_line(raw_line: str) -> StepLineParseResult:
    """Parse one step order line into mark and quantity."""
    line = (raw_line or "").strip()
    if not line:
        return StepLineParseResult(
            parsed=False,
            reason_code="empty_line",
            reason_text="пустая строка",
        )

    mark = extract_step_mark(line)
    if not mark:
        return StepLineParseResult(
            parsed=False,
            reason_code="pattern_not_matched",
            reason_text="не совпал формат строки ступени",
        )

    # Remainder after the mark occurrence in the original line
    match = re.search(r"ЛС\s*\S+", line, re.IGNORECASE | re.UNICODE)
    remainder = line[match.end() :].strip() if match else ""
    # Strip common unit suffixes
    remainder = re.sub(r"\s+шт\.?\b", "", remainder, flags=re.IGNORECASE).strip()
    tokens = remainder.split() if remainder else []

    qty = _parse_qty(tokens)
    if qty is None:
        return StepLineParseResult(
            parsed=False,
            mark=mark,
            reason_code="qty_parse_failed",
            reason_text="не удалось распознать количество",
        )

    return StepLineParseResult(parsed=True, mark=normalize_step_mark(mark), qty=qty)


def merge_step_lines(lines: list[StepLineParseResult]) -> list[StepLineParseResult]:
    """Merge lines with the same mark: sum quantities."""
    merged: dict[str, StepLineParseResult] = {}
    for line in lines:
        if not line.parsed:
            continue
        key = line.mark
        if key in merged:
            existing = merged[key]
            merged[key] = replace(existing, qty=existing.qty + line.qty)
        else:
            merged[key] = replace(line)
    return list(merged.values())


def parse_step_text(text: str) -> list[StepLineParseResult]:
    """Parse multiline step text and merge duplicate marks."""
    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text or "") if part.strip()]
    parsed = [parse_step_line(line) for line in raw_lines]
    return merge_step_lines(parsed)
