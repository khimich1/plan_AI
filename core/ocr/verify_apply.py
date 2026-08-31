#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Choose Extract vs Verify list after the second OCR pass."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable


@dataclass(frozen=True)
class OcrSelectDecision:
    items: list[dict[str, Any]]
    verify_failed: bool
    select_reason: str


def select_ocr_items(
    extract_items: list[dict[str, Any]],
    verify_result: dict[str, Any],
    apply_gate: Callable[[list[dict[str, Any]]], list[dict[str, Any]]],
) -> OcrSelectDecision:
    verified = verify_result.get("plates") or []
    corrections = verify_result.get("corrections") or []
    if not verified:
        return OcrSelectDecision(extract_items, True, "empty_verified_plates")
    if not corrections:
        return OcrSelectDecision(extract_items, False, "kept_extract_empty_corrections")
    return OcrSelectDecision(apply_gate(verified), False, "applied")
