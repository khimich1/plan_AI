#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Локальная проверка распознанных строк через march_line_parser (0 токенов)."""

from __future__ import annotations

from typing import Any, Dict, List

from core.march_line_parser import parse_march_line

_PARSER_REJECTED_ISSUE = "parser_rejected"
_PARSER_REJECTED_CONFIDENCE_CAP = 0.5


def _march_parse_input(march: Dict[str, Any]) -> str:
    candidate = (march.get("normalized_candidate") or march.get("raw_name") or "").strip()
    qty = int(march.get("qty", 1))
    return f"{candidate} {qty}"


def apply_march_parser_gate(marches: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Для каждого марша вызывает parse_march_line; при неуспехе добавляет parser_rejected
    и понижает confidence.
    """
    for march in marches:
        result = parse_march_line(_march_parse_input(march))
        if result.parsed:
            continue

        issues = list(march.get("issues") or [])
        if _PARSER_REJECTED_ISSUE not in issues:
            issues.append(_PARSER_REJECTED_ISSUE)
        march["issues"] = issues

        try:
            confidence = float(march.get("confidence", 0.95))
        except (TypeError, ValueError):
            confidence = 0.95
        march["confidence"] = min(confidence, _PARSER_REJECTED_CONFIDENCE_CAP)

    return marches
