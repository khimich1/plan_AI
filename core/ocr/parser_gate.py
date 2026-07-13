#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Локальная проверка распознанных строк через plate_line_parser (0 токенов)."""

from __future__ import annotations

from typing import Any, Dict, List

from core.plate_line_parser import parse_line

_PARSER_REJECTED_ISSUE = "parser_rejected"
_PARSER_REJECTED_CONFIDENCE_CAP = 0.5


def apply_parser_gate(plates: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Для каждой плиты вызывает parse_line; при неуспехе добавляет parser_rejected
    и понижает confidence.
    """
    for plate in plates:
        candidate = (plate.get("normalized_candidate") or plate.get("raw_name") or "").strip()
        qty = int(plate.get("qty", 1))
        result = parse_line(f"{candidate} {qty}")
        if result.parsed:
            continue

        issues = list(plate.get("issues") or [])
        if _PARSER_REJECTED_ISSUE not in issues:
            issues.append(_PARSER_REJECTED_ISSUE)
        plate["issues"] = issues

        try:
            confidence = float(plate.get("confidence", 0.95))
        except (TypeError, ValueError):
            confidence = 0.95
        plate["confidence"] = min(confidence, _PARSER_REJECTED_CONFIDENCE_CAP)

    return plates
