#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize multiline stair-step (ЛС) order text before parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.step_price_db import extract_step_mark, normalize_step_mark

_DASH_CHARS = "–—‒−"

_STRIP_SHT_RE = re.compile(r"\s+шт\.?\b", re.IGNORECASE | re.UNICODE)

# «Лестничные ступени ЛС11 10» → keep mark + qty
_FULL_NAME_PREFIX_RE = re.compile(
    r"лестничн(?:ые|ая)?\s+ступен[ьяи]\s+",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class StepNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _basic_cleanup(line: str) -> str:
    text = line.replace("\u00a0", " ")
    for ch in _DASH_CHARS:
        text = text.replace(ch, "-")
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def _strip_sht_suffix(line: str) -> str:
    return _STRIP_SHT_RE.sub("", line).strip()


def _normalize_line(line: str) -> str:
    cleaned = _basic_cleanup(line)
    cleaned = _FULL_NAME_PREFIX_RE.sub("", cleaned)
    cleaned = _strip_sht_suffix(cleaned)
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
