#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Марка бетона для позиций заказа и дорожек: КП → pb_reinforcement_series → правило по номеру ПБ."""

from __future__ import annotations

from typing import Any, Mapping, MutableMapping, Optional

from core.db_config import PB_DB_PATH
from core.domain.plate_order import normalize_load_code
from core.reinforcement_db import _extract_length_dm, get_concrete_grade_from_series


def normalize_concrete_grade(raw: Optional[str]) -> Optional[str]:
    """Приводит марку к виду М400 / М500 (кириллическая «М», без пробелов)."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    # Латиница M в начале → кириллица для единообразия отчётов
    if len(s) >= 4 and (s[0] in ("M", "m") and s[1:].strip().startswith(("400", "500"))):
        s = "М" + s[1:].lstrip()
    if s.startswith(("400", "500")):
        s = "М" + s
    if s.startswith("м"):
        s = "М" + s[1:]
    return s if s.startswith("М") else s


def _concrete_grade_by_pb_number(pb_number: int) -> str:
    """До ПБ 59 включительно — М400, с ПБ 60 — М500 (как импорт «по серии»)."""
    if pb_number <= 59:
        return "М400"
    return "М500"


def _pb_number_fallback(plate_name: str, length_m: Optional[float]) -> Optional[int]:
    """Номер ПБ из наименования; при неудаче — из length_m как дециметры (округление)."""
    if plate_name:
        n = _extract_length_dm(plate_name)
        if n is not None:
            return int(n)
    if length_m is None:
        return None
    try:
        return int(round(float(length_m) * 10))
    except Exception:
        return None


def resolve_concrete_grade(
    *,
    concrete_grade_explicit: Optional[str] = None,
    plate_name: str = "",
    length_m: Optional[float] = None,
    load_code: Any = 8,
    db_path: str = PB_DB_PATH,
) -> str:
    """
    Возвращает марку бетона (ненулевая строка).

    Приоритет: явное значение → pb_reinforcement_series → номер ПБ из марки или длины.
    """
    ex = normalize_concrete_grade(concrete_grade_explicit)
    if ex:
        return ex

    lc = normalize_load_code(load_code, default=8)
    if length_m is not None:
        from_series = get_concrete_grade_from_series(float(length_m), lc, db_path=db_path, allow_fallback=True)
        norm_series = normalize_concrete_grade(from_series)
        if norm_series:
            return norm_series

    pb_num = _pb_number_fallback(plate_name or "", length_m)
    if pb_num is not None:
        return _concrete_grade_by_pb_number(pb_num)

    return "М400"


def resolve_concrete_grade_from_order(order: Mapping[str, Any], db_path: str = PB_DB_PATH) -> str:
    """Из словаря orders_2d / позиции плиты (ключи kp_plates-подобные)."""
    lm: float | None = None
    if order.get("length") is not None:
        try:
            lm = float(order["length"])
        except (TypeError, ValueError):
            lm = None
    return resolve_concrete_grade(
        concrete_grade_explicit=order.get("concrete_grade"),
        plate_name=str(order.get("plate_name") or ""),
        length_m=lm,
        load_code=order.get("load_code", 800),
        db_path=db_path,
    )


def enrich_orders_2d_concrete_grade(
    orders_2d: list[MutableMapping[str, Any]],
    *,
    db_path: str = PB_DB_PATH,
) -> None:
    """In-place: для каждой строки задаёт ключ concrete_grade если пуст."""
    for o in orders_2d:
        if normalize_concrete_grade(o.get("concrete_grade")):
            o["concrete_grade"] = normalize_concrete_grade(o.get("concrete_grade"))
            continue
        o["concrete_grade"] = resolve_concrete_grade_from_order(o, db_path=db_path)
