#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Локальная проверка распознанных строк через step_line_parser (0 токенов)."""

from __future__ import annotations

from typing import Any, Dict, List

from core.step_line_parser import parse_step_line

_PARSER_REJECTED_ISSUE = "parser_rejected"
_PARSER_REJECTED_CONFIDENCE_CAP = 0.5


def _step_parse_input(step: Dict[str, Any]) -> str:
    candidate = (step.get("normalized_candidate") or step.get("raw_name") or "").strip()
    qty = int(step.get("qty", 1))
    return f"{candidate} {qty}"


def apply_step_parser_gate(steps: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Для каждой ступени вызывает parse_step_line; при неуспехе понижает confidence."""
    for step in steps:
        result = parse_step_line(_step_parse_input(step))
        if result.parsed:
            continue

        issues = list(step.get("issues") or [])
        if _PARSER_REJECTED_ISSUE not in issues:
            issues.append(_PARSER_REJECTED_ISSUE)
        step["issues"] = issues

        try:
            confidence = float(step.get("confidence", 0.95))
        except (TypeError, ValueError):
            confidence = 0.95
        step["confidence"] = min(confidence, _PARSER_REJECTED_CONFIDENCE_CAP)

    return steps
