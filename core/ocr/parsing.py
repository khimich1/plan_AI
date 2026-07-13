#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Парсинг JSON-ответов OCR (Extract и Verify)."""

from __future__ import annotations

import json
import re
from typing import Any, Dict, List, Optional


def _validate_plate_item(plate: Any) -> Optional[Dict[str, Any]]:
    if not isinstance(plate, dict):
        print(f"[GPT] ⚠️ Пропущена плита (не объект): {plate}")
        return None

    raw_name = plate.get("raw_name") or plate.get("name")
    normalized_candidate = plate.get("normalized_candidate") or raw_name
    if not raw_name or "qty" not in plate:
        print(f"[GPT] ⚠️ Пропущена плита (нет raw_name/name или qty): {plate}")
        return None

    try:
        qty = int(plate["qty"])
    except (ValueError, TypeError):
        print(f"[GPT] ⚠️ Пропущена плита (qty не число): {plate}")
        return None

    confidence = plate.get("confidence", 0.95)
    try:
        confidence = max(0.0, min(1.0, float(confidence)))
    except (ValueError, TypeError):
        confidence = 0.95

    issues = plate.get("issues") if isinstance(plate.get("issues"), list) else []
    return {
        "raw_name": str(raw_name).strip(),
        "normalized_candidate": str(normalized_candidate).strip(),
        "qty": qty,
        "confidence": confidence,
        "issues": issues,
    }


def _extract_json_from_response(response_text: str) -> str:
    text = (response_text or "").strip()
    object_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if object_match:
        return object_match.group(1)
    array_match = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", text, re.DOTALL)
    if array_match:
        return array_match.group(1)
    if text.startswith("{") or text.startswith("["):
        return text
    brace_match = re.search(r"(\{.*\})", text, re.DOTALL)
    if brace_match:
        return brace_match.group(1)
    bracket_match = re.search(r"(\[.*\])", text, re.DOTALL)
    if bracket_match:
        return bracket_match.group(1)
    return text


def parse_gpt_response(response_text: str) -> List[Dict[str, Any]]:
    """Извлекает JSON-массив плит из ответа GPT (этап Extract)."""
    json_text = _extract_json_from_response(response_text)

    try:
        parsed = json.loads(json_text)
        if isinstance(parsed, dict) and "plates" in parsed:
            parsed = parsed["plates"]
        if not isinstance(parsed, list):
            print(f"[GPT] ⚠️ Ожидался JSON-массив, получено: {type(parsed)}")
            return []

        validated_plates = []
        for plate in parsed:
            item = _validate_plate_item(plate)
            if item:
                validated_plates.append(item)
        return validated_plates

    except json.JSONDecodeError as e:
        print(f"[GPT] ❌ Ошибка парсинга JSON: {e}")
        print(f"[GPT] Ответ GPT (первые 200 символов):")
        print(response_text[:200])
        return []


def parse_verify_response(response_text: str) -> Dict[str, Any]:
    """
    Парсит ответ этапа Verify.
    Поддерживает полный объект {{plates, corrections}} и legacy-массив plates.
    """
    json_text = _extract_json_from_response(response_text)

    try:
        parsed = json.loads(json_text)

        if isinstance(parsed, list):
            plates_raw = parsed
            corrections: List[Dict[str, Any]] = []
            row_count_on_image = len(plates_raw)
        elif isinstance(parsed, dict):
            plates_raw = parsed.get("plates") or []
            corrections = parsed.get("corrections") if isinstance(parsed.get("corrections"), list) else []
            row_count_raw = parsed.get("row_count_on_image")
            try:
                row_count_on_image = int(row_count_raw) if row_count_raw is not None else None
            except (ValueError, TypeError):
                row_count_on_image = None
        else:
            print(f"[GPT] ⚠️ Verify: неожиданный тип JSON: {type(parsed)}")
            return {"plates": [], "corrections": [], "row_count_on_image": None}

        validated_plates = []
        for plate in plates_raw:
            item = _validate_plate_item(plate)
            if item:
                validated_plates.append(item)

        return {
            "plates": validated_plates,
            "corrections": corrections,
            "row_count_on_image": row_count_on_image,
        }

    except json.JSONDecodeError as e:
        print(f"[GPT] ❌ Verify: ошибка парсинга JSON: {e}")
        print(f"[GPT] Ответ GPT (первые 200 символов):")
        print(response_text[:200])
        return {"plates": [], "corrections": [], "row_count_on_image": None}
