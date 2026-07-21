#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Детерминированный сплит строк с маркером добора (+доб) на две канонические позиции.

Формула: W_доб = 12 − W (в дм). Пример:
  «ПБ 57-7,2-8п + доб 5-шт» → «ПБ 57-7,2-8п 5» и «ПБ 57-4,8-8п 5»
"""
from __future__ import annotations

import re
from dataclasses import dataclass

_FORM_WIDTH_DM = 12.0

_DOБOR_MARKER_RE = re.compile(
    r"(?i)"
    r"\s*(?:\+\s*)?"
    r"(?:доб(?:ор)?)"
    r"(?:\s*[-—]?\s*)?"
    r"(?:(\d+)(?:\s*[-—]?\s*)?(?:шт\.?)?)?"
    r"\s*$",
    re.UNICODE,
)

_PLATE_BEFORE_DOBOR_RE = re.compile(
    r"(?i)"
    r"^\s*"
    r"(п[бк])"
    r"\s+"
    r"([\d\.,]+)"
    r"\s*-\s*"
    r"([\d\.,]+)"
    r"\s*-\s*"
    r"([\d\.,]+)"
    r"п?"
    r"(?:\s+(\d+)\s*(?:шт\.?)?)?"
    r"\s*$",
    re.UNICODE,
)


@dataclass(frozen=True)
class DoborPair:
    pair_id: str
    primary_line: str
    complement_line: str
    source_line: str


def _width_dm_to_float(wdm_str: str) -> float:
    s = (wdm_str or "").strip().replace(",", ".")
    try:
        val = float(s)
    except ValueError:
        return 0.0
    if "." in s and val <= 2.0:
        return val * 10.0
    return val


def _format_width_dm(W_dm: float) -> str:
    rounded = round(W_dm, 1)
    if abs(rounded - round(rounded)) < 1e-6:
        if rounded < 10:
            return f"{int(round(rounded))},0"
        return str(int(round(rounded)))
    return f"{rounded:.1f}".replace(".", ",")


def _format_load(load: float) -> str:
    if abs(load - round(load)) < 1e-6:
        return f"{int(round(load))}п"
    return f"{load:.1f}п".replace(".", ",")


def _build_canonical_line_from_parts(
    prefix: str,
    ldm_str: str,
    width_dm: float,
    load: float,
    qty: int,
) -> str:
    w_str = _format_width_dm(width_dm)
    canonical = f"{prefix.upper()} {ldm_str.strip()}-{w_str}-{_format_load(load)}"
    if qty > 1:
        canonical += f" {qty}"
    return canonical


def expand_dobor_line(line: str, *, pair_index: int = 1) -> tuple[list[str], DoborPair | None, list[str]]:
    """
    Разворачивает строку с маркером добора в две канонические позиции.

    «ПБ 57-7,2-8п + доб 5-шт» →
      (["ПБ 57-7,2-8п 5", "ПБ 57-4,8-8п 5"], DoborPair(...), [])

    Без маркера добора → ([canonical_or_cleaned], None, [])
    """
    warnings: list[str] = []
    stripped = (line or "").strip()
    if not stripped:
        return [stripped], None, warnings

    from .plate_text_normalizer import basic_text_cleanup, normalize_plate_prefixes

    cleaned = normalize_plate_prefixes(basic_text_cleanup(stripped))

    marker_match = _DOБOR_MARKER_RE.search(cleaned)
    if not marker_match:
        return [cleaned], None, warnings

    dobor_qty_raw = marker_match.group(1)
    dobor_qty = int(dobor_qty_raw) if dobor_qty_raw else None
    plate_part = cleaned[: marker_match.start()].strip()

    plate_match = _PLATE_BEFORE_DOBOR_RE.match(plate_part)
    if not plate_match:
        return [cleaned], None, warnings

    raw_prefix = plate_match.group(1)
    ldm_str = plate_match.group(2)
    wdm_str = plate_match.group(3)
    load_str = plate_match.group(4)
    plate_qty_raw = plate_match.group(5)

    try:
        load = float(load_str.replace(",", "."))
    except ValueError:
        load = 8.0

    plate_qty = int(plate_qty_raw) if plate_qty_raw else None

    if dobor_qty is not None:
        qty = max(1, dobor_qty)
        if plate_qty is not None and plate_qty != qty:
            warnings.append(
                f"Конфликт количества в строке {stripped!r}: у марки {plate_qty}, у добора {dobor_qty}; "
                f"использовано количество добора ({qty})"
            )
    elif plate_qty is not None:
        qty = max(1, plate_qty)
    else:
        qty = 1

    w_dm = _width_dm_to_float(wdm_str)
    if w_dm <= 0:
        return [cleaned], None, warnings

    complement_dm = round(_FORM_WIDTH_DM - w_dm, 1)
    if w_dm >= _FORM_WIDTH_DM or complement_dm <= 0:
        warnings.append(
            f"Добор невозможен для ширины {wdm_str} дм в строке {stripped!r}: "
            f"остаток {complement_dm} дм; строка оставлена без сплита"
        )
        return [cleaned], None, warnings

    primary_line = _build_canonical_line_from_parts(raw_prefix, ldm_str, w_dm, load, qty)
    complement_line = _build_canonical_line_from_parts(raw_prefix, ldm_str, complement_dm, load, qty)

    pair_id = f"dobor-{pair_index}"
    pair = DoborPair(
        pair_id=pair_id,
        primary_line=primary_line,
        complement_line=complement_line,
        source_line=stripped,
    )
    return [primary_line, complement_line], pair, warnings
