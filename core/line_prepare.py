#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Per-line prepare of commercial source text before parse_*_line.

Shared format cleanup for all product types. GOST→factory rewrite only for
bridge piles with an explicit class (T/В + number). Does not rewrite the
textarea; callers parse a copy.
"""

from __future__ import annotations

import re

_DASH_CHARS = "–—‒−"

_LIST_PREFIX_RE = re.compile(r"^\d+[.),]\s+")
# Qty unit after a digit; not from the middle of a mark (`C14шт35`).
_QTY_UNIT_RE = re.compile(r"(?<=\d)\s*шт(?:ук)?\.?(?=\s|$)", re.IGNORECASE | re.UNICODE)

# Factory: C14-35T7 / С7-35Т5 / C8-35В4
_FACTORY_BRIDGE_MARK_RE = re.compile(
    r"^[СC]\s*\d+\s*-\s*\d+\s*[TТBВ]\s*\d+",
    re.IGNORECASE | re.UNICODE,
)
# GOST: C 14.35-T7 / С14,35-В4 — length .|, section - class+number
_GOST_BRIDGE_MARK_RE = re.compile(
    r"^([СC])\s*(\d+)\s*[.,]\s*(\d+)\s*-\s*([TТBВ])\s*(\d+)",
    re.IGNORECASE | re.UNICODE,
)

_BRIDGE_PRODUCT_TYPE = "bridge_piles"


def cleanup_source_line(raw: str) -> str:
    """Shared format cleanup: NBSP/dashes, list numbering, qty units, whitespace."""
    text = (raw or "").replace("\u00a0", " ")
    for ch in _DASH_CHARS:
        text = text.replace(ch, "-")
    text = _LIST_PREFIX_RE.sub("", text)
    text = _QTY_UNIT_RE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _rewrite_gost_bridge_prefix(text: str) -> str:
    if not text or _FACTORY_BRIDGE_MARK_RE.match(text):
        return text
    match = _GOST_BRIDGE_MARK_RE.match(text)
    if not match:
        return text
    letter_c, length, section, class_letter, class_num = match.groups()
    factory = f"{letter_c}{length}-{section}{class_letter}{class_num}"
    remainder = text[match.end() :].strip()
    if remainder:
        return f"{factory} {remainder}"
    return factory


def prepare_bridge_pile_mark(mark: str) -> str:
    """GOST→factory rewrite of a bridge-pile mark token. Idempotent on factory marks."""
    return _rewrite_gost_bridge_prefix(cleanup_source_line(mark))


def prepare_source_line(raw: str, product_type: str) -> str:
    """Copy of ``raw`` for the parser: shared cleanup, plus bridge GOST rewrite."""
    cleaned = cleanup_source_line(raw)
    if product_type != _BRIDGE_PRODUCT_TYPE:
        return cleaned
    return _rewrite_gost_bridge_prefix(cleaned)


def prepare_candidate_qty_line(candidate: str, qty: int, product_type: str) -> str:
    """Prepare OCR ``candidate + qty`` without doubling qty already in the candidate."""
    prepared_candidate = prepare_source_line(candidate, product_type)
    tokens = prepared_candidate.split()
    if tokens and tokens[-1] == str(int(qty)):
        return prepared_candidate
    return prepare_source_line(f"{candidate} {qty}", product_type)
