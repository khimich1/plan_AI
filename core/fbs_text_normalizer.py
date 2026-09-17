#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize multiline FBS order text before parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.line_prepare import prepare_source_line


@dataclass
class FbsNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _normalize_line(line: str) -> str:
    return prepare_source_line(line, "fbs")


def normalize_fbs_order_text(text: str) -> FbsNormalizeResult:
    """Split multiline FBS text; light whitespace/dash cleanup only."""
    if not text or not text.strip():
        return FbsNormalizeResult(normalized_text=text or "")

    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text) if part.strip()]
    normalized_lines = [_normalize_line(line) for line in raw_lines]
    return FbsNormalizeResult(
        normalized_text="\n".join(normalized_lines),
        normalized_lines=normalized_lines,
    )
