#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Локальная проверка распознанных строк через pile_line_parser (0 токенов)."""

from __future__ import annotations

from typing import Any, Dict, List

from core.pile_line_parser import parse_pile_line

_PARSER_REJECTED_ISSUE = "parser_rejected"
_PARSER_REJECTED_CONFIDENCE_CAP = 0.5


def _pile_parse_input(pile: Dict[str, Any]) -> str:
    candidate = (pile.get("normalized_candidate") or pile.get("raw_name") or "").strip()
    qty = int(pile.get("qty", 1))
    return f"{candidate} {qty}"


def apply_pile_parser_gate(piles: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Для каждой сваи вызывает parse_pile_line; при неуспехе добавляет parser_rejected
    и понижает confidence.
    """
    for pile in piles:
        result = parse_pile_line(_pile_parse_input(pile))
        if result.parsed:
            continue

        issues = list(pile.get("issues") or [])
        if _PARSER_REJECTED_ISSUE not in issues:
            issues.append(_PARSER_REJECTED_ISSUE)
        pile["issues"] = issues

        try:
            confidence = float(pile.get("confidence", 0.95))
        except (TypeError, ValueError):
            confidence = 0.95
        pile["confidence"] = min(confidence, _PARSER_REJECTED_CONFIDENCE_CAP)

    return piles
