#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize multiline stair-step (ЛС) order text before parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.step_price_db import extract_step_mark, normalize_step_mark
from core.line_prepare import prepare_source_line

# «Лестничные ступени ЛС11 10» → keep mark + qty
_FULL_NAME_PREFIX_RE = re.compile(
    r"лестничн(?:ые|ая)?\s+ступен[ьяи]\s+",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class StepNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _normalize_line(line: str) -> str:
    cleaned = prepare_source_line(line, "steps")
    cleaned = _FULL_NAME_PREFIX_RE.sub("", cleaned)
    mark = extract_step_mark(cleaned)
    if mark:
        # Rebuild as «MARK qty…» with normalized mark
        match = re.search(r"ЛС\s*\S+", cleaned, re.IGNORECASE | re.UNICODE)
        remainder = cleaned[match.end() :].strip() if match else ""
        if remainder:
            return f"{normalize_step_mark(mark)} {remainder}".strip()
        return normalize_step_mark(mark)
    return cleaned.strip()


def normalize_step_order_text(text: str) -> StepNormalizeResult:
    """Split multiline step text and apply whitespace/dash/full-name normalization."""
    if not text or not text.strip():
        return StepNormalizeResult(normalized_text=text or "")

    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text) if part.strip()]
    normalized_lines = [_normalize_line(line) for line in raw_lines]
    return StepNormalizeResult(
        normalized_text="\n".join(normalized_lines),
        normalized_lines=normalized_lines,
    )
