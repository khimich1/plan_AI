#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize multiline bridge-pile order text before parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.line_prepare import prepare_source_line


@dataclass
class BridgePileNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _normalize_line(line: str) -> str:
    return prepare_source_line(line, "bridge_piles")


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
