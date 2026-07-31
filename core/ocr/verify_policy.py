#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Решение о втором API-вызове (Verify) по режиму и эвристикам auto."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Tuple

from core.pile_line_parser import parse_pile_line
from core.plate_line_parser import parse_line


@dataclass(frozen=True)
class OcrVerifySettings:
    """Пороги режима auto (из env / settings)."""

    max_rows: int = 10
    min_confidence: float = 0.92
    max_bytes: int = 819_200


def _plate_parse_line(plate: Dict[str, Any]) -> str:
    candidate = (plate.get("normalized_candidate") or plate.get("raw_name") or "").strip()
    qty = int(plate.get("qty", 1))
    return f"{candidate} {qty}"


def _plate_confidence(plate: Dict[str, Any]) -> float:
    try:
        return float(plate.get("confidence", 0.95))
    except (TypeError, ValueError):
        return 0.95


def should_run_verify(
    *,
    mode: str,
    max_api_calls: int,
    image_size_bytes: int,
    plates: List[Dict[str, Any]],
    settings: OcrVerifySettings,
) -> Tuple[bool, str]:
    """
    Возвращает (run_verify, reason).

    reason пишется в metadata: ocr_verify_skipped_reason или ocr_verify_applied_reason.
    """
    if max_api_calls <= 1 or mode == "never":
        return False, "max_api_calls_or_never"

    if mode == "always":
        return True, "mode_always"

    if mode != "auto":
        return True, f"unknown_mode_{mode}"

    if not plates:
        return True, "auto_empty_plates"

    if image_size_bytes > settings.max_bytes:
        return True, "auto_file_too_large"

    if len(plates) > settings.max_rows:
        return True, "auto_too_many_rows"

    for plate in plates:
        if _plate_confidence(plate) < settings.min_confidence:
            return True, "auto_low_confidence"

        if not parse_line(_plate_parse_line(plate)).parsed:
            return True, "auto_unparsed_plate"

        issues = plate.get("issues")
        if issues:
            return True, "auto_has_issues"

    return False, "auto_all_checks_passed"


def _pile_parse_input(pile: Dict[str, Any]) -> str:
    candidate = (pile.get("normalized_candidate") or pile.get("raw_name") or "").strip()
    qty = int(pile.get("qty", 1))
    return f"{candidate} {qty}"


def _pile_confidence(pile: Dict[str, Any]) -> float:
    try:
        return float(pile.get("confidence", 0.95))
    except (TypeError, ValueError):
        return 0.95


def should_run_pile_verify(
    *,
    mode: str,
    max_api_calls: int,
    image_size_bytes: int,
    piles: List[Dict[str, Any]],
    settings: OcrVerifySettings,
) -> Tuple[bool, str]:
    """
    Возвращает (run_verify, reason) для OCR свай.

    reason пишется в metadata: ocr_verify_skipped_reason или ocr_verify_applied_reason.
    """
    if max_api_calls <= 1 or mode == "never":
        return False, "max_api_calls_or_never"

    if mode == "always":
        return True, "mode_always"

    if mode != "auto":
        return True, f"unknown_mode_{mode}"

    if not piles:
        return True, "auto_empty_piles"

    if image_size_bytes > settings.max_bytes:
        return True, "auto_file_too_large"

    if len(piles) > settings.max_rows:
        return True, "auto_too_many_rows"

    for pile in piles:
        if _pile_confidence(pile) < settings.min_confidence:
            return True, "auto_low_confidence"

        if not parse_pile_line(_pile_parse_input(pile)).parsed:
            return True, "auto_unparsed_pile"

        issues = pile.get("issues")
        if issues:
            return True, "auto_has_issues"

    return False, "auto_all_checks_passed"
