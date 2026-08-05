#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize multiline bridge-pile order text before parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_DASH_CHARS = "–—‒−"

_STRIP_SHT_RE = re.compile(r"\s+шт\.?\b", re.IGNORECASE | re.UNICODE)


@dataclass
class BridgePileNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _basic_cleanup(line: str) -> str:
    text = line.replace("\u00a0", " ")
    for ch in _DASH_CHARS:
        text = text.replace(ch, "-")
    text = re.sub(r"\s{2,}", " ", text)
    text = _STRIP_SHT_RE.sub("", text).strip()
    return text.strip()


def _normalize_line(line: str) -> str:
    return _basic_cleanup(line)


def normalize_bridge_pile_order_text(text: str) -> BridgePileNormalizeResult:
    """Split multiline bridge-pile text; light whitespace/dash cleanup only."""
    if not text or not text.strip():
        return BridgePileNormalizeResult(normalized_text=text or "")

    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text) if part.strip()]
    normalized_lines = [_normalize_line(line) for line in raw_lines]
    return BridgePileNormalizeResult(
        normalized_text="\n".join(normalized_lines),
        normalized_lines=normalized_lines,
    )
