#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize multiline pile order text before parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_DASH_CHARS = "–—‒−"

# «Сваи 90.30-11 189» → «С90.30-11 189» (GPT оторвал С от марки).
_REPAIR_SVAI_PREFIX_RE = re.compile(
    r"\b[СC]ваи\s+(\d[\d.,]*(?:-[\d.,]+(?:[иИ])?)?)",
    re.IGNORECASE | re.UNICODE,
)

# «Свай 110.30-13 26» → «С110.30-13 26» (единственное число в OCR).
_REPAIR_SVAY_PREFIX_RE = re.compile(
    r"\b[СC]вай\s+(\d[\d.,]*(?:-[\d.,]+(?:[иИ])?)?)",
    re.IGNORECASE | re.UNICODE,
)

# «… 189 шт» → «… 189»
_STRIP_SHT_RE = re.compile(r"\s+шт\.?\b", re.IGNORECASE | re.UNICODE)

# OCR: «С120.30-12 20» — 20 в колонке qty, не класс B20 (только после R1/R2).
_OCR_QTY_GRADE_AMBIGUOUS = frozenset({"15", "20", "22.5", "22,5", "25", "30"})
_MARK_WITH_TRAILING_TOKEN_RE = re.compile(
    r"^([СC][\d.,]+(?:-[\d.,]+(?:[иИ])?)?)\s+(\S+)\s*$",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class PileNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _basic_cleanup(line: str) -> str:
    text = line.replace("\u00a0", " ")
    for ch in _DASH_CHARS:
        text = text.replace(ch, "-")
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def _repair_missing_cyrillic_c_mark(line: str) -> str:
    """«Сваи 90.30-11 189» → «С90.30-11 189» (GPT оторвал С от марки)."""
    return _REPAIR_SVAI_PREFIX_RE.sub(r"С\1", line)


def _repair_svay_singular_prefix(line: str) -> str:
    """«Свай 110.30-13 26» → «С110.30-13 26»."""
    return _REPAIR_SVAY_PREFIX_RE.sub(r"С\1", line)


def _strip_sht_suffix(line: str) -> str:
    """«… 189 шт» → «… 189»."""
    return _STRIP_SHT_RE.sub("", line).strip()


def _disambiguate_ocr_qty_vs_grade(line: str) -> str:
    """После OCR-префикса «Сваи/Свай»: «С120.30-12 20» → «С120.30-12 B25 20»."""
    match = _MARK_WITH_TRAILING_TOKEN_RE.match(line)
    if not match:
        return line
    mark = re.sub(r"^[СC]\s*", "С", match.group(1), flags=re.IGNORECASE)
    token = match.group(2).strip().lower().replace(",", ".")
    if token not in _OCR_QTY_GRADE_AMBIGUOUS:
        return line
    return f"{mark} B25 {match.group(2).strip()}"


def _normalize_line(line: str) -> str:
    cleaned = _basic_cleanup(line)
    had_ocr_prefix = bool(
        _REPAIR_SVAI_PREFIX_RE.search(cleaned) or _REPAIR_SVAY_PREFIX_RE.search(cleaned)
    )
    cleaned = _repair_missing_cyrillic_c_mark(cleaned)
    cleaned = _repair_svay_singular_prefix(cleaned)
    cleaned = _strip_sht_suffix(cleaned)
    if had_ocr_prefix:
        cleaned = _disambiguate_ocr_qty_vs_grade(cleaned)
    cleaned = re.sub(r"^[СC]\s+", "С", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


def normalize_pile_order_text(text: str) -> PileNormalizeResult:
    """Split multiline pile text and apply OCR repair + whitespace/dash normalization."""
    if not text or not text.strip():
        return PileNormalizeResult(normalized_text=text or "")

    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text) if part.strip()]
    normalized_lines = [_normalize_line(line) for line in raw_lines]
    return PileNormalizeResult(
        normalized_text="\n".join(normalized_lines),
        normalized_lines=normalized_lines,
    )
