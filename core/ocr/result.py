#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка результата OCR и форматирование для UI."""

from __future__ import annotations

from typing import Any, Dict, List, Optional


def plates_to_text(plates: List[Dict[str, Any]]) -> str:
    text_lines: List[str] = []
    for plate in plates:
        candidate = (plate.get("normalized_candidate") or plate.get("raw_name") or "").strip()
        qty = int(plate.get("qty", 1))
        if candidate:
            text_lines.append(f"{candidate} {qty}")
    return "\n".join(text_lines)


def build_result_payload(
    *,
    plates: List[Dict[str, Any]],
    draft_plates: List[Dict[str, Any]],
    corrections: List[Dict[str, Any]],
    row_count_on_image: Optional[int],
    method: str,
    verify_applied: bool,
    verify_failed: bool,
    cost_usd: float,
    cost_rub: float = 0.0,
    api_calls: int = 1,
    verify_skipped_reason: Optional[str] = None,
    verify_applied_reason: Optional[str] = None,
    verify_select_reason: Optional[str] = None,
    ocr_preprocess: Optional[str] = None,
) -> Dict[str, Any]:
    confidences = [float(p.get("confidence", 0.95)) for p in plates]
    avg_confidence = sum(confidences) / len(confidences) if confidences else 0.95
    return {
        "text": plates_to_text(plates),
        "plates": plates,
        "draft_plates": draft_plates,
        "corrections": corrections,
        "row_count_on_image": row_count_on_image,
        "method": method,
        "verify_applied": verify_applied,
        "verify_failed": verify_failed,
        "confidence": avg_confidence,
        "cost_usd": cost_usd,
        "ocr_api_calls": api_calls,
        "ocr_cost_rub": cost_rub,
        "ocr_cost_usd": cost_usd,
        "ocr_method": method,
        "ocr_verify_skipped_reason": verify_skipped_reason,
        "ocr_verify_applied_reason": verify_applied_reason,
        "ocr_verify_select_reason": verify_select_reason,
        "ocr_preprocess": ocr_preprocess,
    }


def format_corrections_for_user(
    corrections: List[Dict[str, Any]],
    *,
    max_items: int = 8,
) -> str:
    """Краткий текст исправлений для Telegram."""
    actionable = [
        c for c in corrections
        if c.get("action") != "verify_failed"
    ]
    if not actionable:
        return ""

    lines = [f"⚠️ Автоисправлено {len(actionable)} строк(и):"]
    for idx, item in enumerate(actionable[:max_items], start=1):
        action = item.get("action") or "changed"
        row_index = item.get("row_index")
        row_label = f"стр. {row_index}" if row_index is not None else f"#{idx}"

        before = item.get("before") or {}
        after = item.get("after") or {}
        before_mark = before.get("normalized_candidate") or before.get("raw_name") or "—"
        after_mark = after.get("normalized_candidate") or after.get("raw_name") or "—"
        before_qty = before.get("qty")
        after_qty = after.get("qty")

        if action == "added":
            mark = after_mark
            qty = after_qty if after_qty is not None else "?"
            lines.append(f"• {row_label}: добавлено «{mark} {qty}»")
        elif action == "removed":
            lines.append(f"• {row_label}: удалено «{before_mark}»")
        elif action == "changed_qty":
            lines.append(
                f"• {row_label}: «{after_mark}» qty {before_qty} → {after_qty}"
            )
        elif action == "changed_mark":
            lines.append(
                f"• {row_label}: «{before_mark}» → «{after_mark}»"
            )
        elif action == "reordered":
            lines.append(f"• {row_label}: изменён порядок")
        else:
            reason = item.get("reason") or action
            lines.append(f"• {row_label}: {reason}")

    if len(actionable) > max_items:
        lines.append(f"• … и ещё {len(actionable) - max_items}")

    return "\n".join(lines)


def estimate_monthly_cost(photos_per_month: int) -> Dict[str, float]:
    """Оценка месячных затрат на OCR (один вызов GPT-4o)."""
    avg_cost_per_photo = 0.002

    return {
        "gpt_only": photos_per_month * avg_cost_per_photo,
        "hybrid": photos_per_month * avg_cost_per_photo,
        "photos": photos_per_month,
    }
