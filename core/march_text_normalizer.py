#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize multiline stair-march (ЛМ) order text before parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.march_price_db import normalize_march_mark
from core.line_prepare import prepare_source_line

_FULL_NAME_PREFIX_RE = re.compile(
    r"лестничн(?:ые|ая|ый)?\s+марш[иае]?\s+",
    re.IGNORECASE | re.UNICODE,
)

# Mark at start: 1ЛМ … or ЛМ 2,8 / ЛМ 2.8 (+ optional закладные справа)
_MARK_AT_START_RE = re.compile(
    r"^("
    r"1ЛМ\s*\d+(?:\s*-\s*\d+)+(?:\s+закладные\s+справа)?|"
    r"ЛМ\s*\d+[.,]\d+"
    r")",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class MarchNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _normalize_line(line: str) -> str:
    cleaned = prepare_source_line(line, "marches")
    cleaned = _FULL_NAME_PREFIX_RE.sub("", cleaned)
    match = _MARK_AT_START_RE.match(cleaned)
    if match:
        raw_mark = match.group(1)
        remainder = cleaned[match.end() :].strip()
        mark = normalize_march_mark(raw_mark)
        if remainder:
            return f"{mark} {remainder}".strip()
        return mark
    return cleaned.strip()


def normalize_march_order_text(text: str) -> MarchNormalizeResult:
    """Split multiline march text and apply whitespace/dash/full-name normalization."""
    if not text or not text.strip():
        return MarchNormalizeResult(normalized_text=text or "")

    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text) if part.strip()]
    normalized_lines = [_normalize_line(line) for line in raw_lines]
    return MarchNormalizeResult(
        normalized_text="\n".join(normalized_lines),
        normalized_lines=normalized_lines,
    )
