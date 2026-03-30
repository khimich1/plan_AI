#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Разбор одной строки заказа в структурированный результат.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional


@dataclass
class LineParseResult:
    parsed: bool
    stage: str
    width_m: float = 0.0
    length_m: float = 0.0
    qty: int = 1
    load_code: Optional[float] = None
    length_dm_raw: str = ""
    reason_code: str = ""
    reason_text: str = ""


_WXL_RE = re.compile(r"(\d+(?:\.\d+)?)\s*[xх]\s*(\d+(?:\.\d+)?)\D*(\d+)?", re.IGNORECASE)
_PLATE_MARK_RE = re.compile(
    r"(?:плит[аы]?\s*)?\bп[бк][\s\.,]*([\d\.,]+)\s*-\s*([\d\.,]+)",
    re.IGNORECASE,
)
_QTY_AFTER_LOAD_DASH_RE = re.compile(r"-\d+п\s*[-—–]\s*(\d+)", re.IGNORECASE)
_QTY_AFTER_LOAD_RE = re.compile(r"-\d+п\s+(\d+)\s*(шт)?\s*$", re.IGNORECASE)
_QTY_END_RE = re.compile(r"(?<!-\d)(\d+)\s*(шт)?\s*$", re.IGNORECASE)
_LOAD_WITH_P_RE = re.compile(r"-\s*([\d\.,]+)\s*п\b", re.IGNORECASE)
_LOAD_PLAIN_RE = re.compile(
    r"п[бк][\s\.,]*[\d\.,]+\s*-\s*[\d\.,]+\s*-\s*([\d\.,]+)",
    re.IGNORECASE,
)


def _length_dm_to_m(ldm_str: str) -> float:
    s = (ldm_str or "").strip().replace(" ", "")
    if "," in s or "." in s:
        parts = s.replace(",", ".").split(".")
        if len(parts) != 2:
            try:
                return round(float(s.replace(",", ".")) / 10.0, 3)
            except ValueError:
                return 0.0
        try:
            int_part = int(parts[0].strip())
            frac_str = (parts[1].strip() or "0")[:3]
            frac_part = int(frac_str) if frac_str else 0
            denom = 10 ** len(frac_str)
            length_mm = (int_part * denom + frac_part) * 100 // denom
        except ValueError:
            return 0.0
        if length_mm < 0:
            return 0.0
        return round(length_mm / 1000.0, 3)
    try:
        nominal_dm = int(float(s))
    except ValueError:
        return 0.0
    return round(nominal_dm / 10.0, 3)


def _normalize_dimension(value_str: str) -> float:
    clean = value_str.strip().replace(" ", "").replace(",", ".")
    try:
        val = float(clean)
    except ValueError:
        return 0.0
    if 20.0 < val < 1000.0:
        val = val / 100.0
    return val


def _parse_pb_width_to_m(width_str: str) -> float:
    width_raw = _normalize_dimension(width_str)
    if width_raw <= 0:
        return 0.0
    if abs(width_raw - 0.3) < 1e-6 or abs(width_raw - 0.2) < 1e-6:
        return round(width_raw, 3)
    return round(width_raw / 10.0, 3)


def parse_line(raw_line: str) -> LineParseResult:
    """
    Возвращает структурированный результат разбора строки.
    """
    s = (raw_line or "").strip()
    if not s:
        return LineParseResult(False, "empty", reason_code="empty_line", reason_text="пустая строка")

    s_lower = s.lower()
    s_norm = s_lower.replace(",", ".")

    # 1) WxL
    m_wxl = _WXL_RE.search(s_norm)
    if m_wxl:
        first = float(m_wxl.group(1))
        second = float(m_wxl.group(2))
        qty = int((m_wxl.group(3) or "1").replace(",", "."))
        if first > 2.0 and second <= 1.5:
            width_m, length_m = second, first
        else:
            width_m, length_m = first, second
        return LineParseResult(
            parsed=True,
            stage="strict_wxl",
            width_m=round(width_m, 3),
            length_m=round(length_m, 3),
            qty=qty,
        )

    # 2) Марка ПБ/ПК
    m_mark = _PLATE_MARK_RE.search(s_lower)
    if not m_mark:
        return LineParseResult(False, "unparsed", reason_code="pattern_not_matched", reason_text="не совпал формат строки")

    ldm_str = m_mark.group(1)
    wdm_str = m_mark.group(2)
    length_m = _length_dm_to_m(ldm_str)
    width_m = _parse_pb_width_to_m(wdm_str)
    if length_m <= 0 or width_m <= 0:
        return LineParseResult(
            False,
            "tolerant_pbpk",
            reason_code="dimension_parse_failed",
            reason_text="не удалось распознать длину/ширину",
        )

    qty = 1
    qty_match = _QTY_AFTER_LOAD_DASH_RE.search(s_lower) or _QTY_AFTER_LOAD_RE.search(s_lower) or _QTY_END_RE.search(s_lower)
    if qty_match:
        try:
            qty = int(qty_match.group(1))
        except Exception:
            qty = 1

    load_code: Optional[float] = None
    load_match = _LOAD_WITH_P_RE.search(s_lower) or _LOAD_PLAIN_RE.search(s_lower)
    if load_match:
        try:
            load_val = float(load_match.group(1).replace(",", "."))
            if load_val > 0:
                load_code = load_val
        except Exception:
            load_code = None

    return LineParseResult(
        parsed=True,
        stage="tolerant_pbpk",
        width_m=round(width_m, 3),
        length_m=round(length_m, 3),
        qty=qty,
        load_code=load_code,
        length_dm_raw=ldm_str.strip() if ldm_str else "",
    )
