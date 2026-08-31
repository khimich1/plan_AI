#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Локальная проверка распознанных строк через fbs_line_parser."""

from __future__ import annotations

from typing import Any, Dict, List

from core.fbs_line_parser import parse_fbs_line

_PARSER_REJECTED_ISSUE = "parser_rejected"
_PARSER_REJECTED_CONFIDENCE_CAP = 0.5


def _fbs_parse_input(item: Dict[str, Any]) -> str:
    candidate = (item.get("normalized_candidate") or item.get("raw_name") or "").strip()
    qty = int(item.get("qty", 1))
    return f"{candidate} {qty}"


def apply_fbs_parser_gate(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    for item in items:
        result = parse_fbs_line(_fbs_parse_input(item))
        if result.parsed:
            continue

        issues = list(item.get("issues") or [])
        if _PARSER_REJECTED_ISSUE not in issues:
            issues.append(_PARSER_REJECTED_ISSUE)
        item["issues"] = issues

        try:
            confidence = float(item.get("confidence", 0.95))
        except (TypeError, ValueError):
            confidence = 0.95
        item["confidence"] = min(confidence, _PARSER_REJECTED_CONFIDENCE_CAP)

    return items
