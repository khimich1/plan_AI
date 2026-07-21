#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Маскирование чувствительных полей в логах (OPT-006)."""

from __future__ import annotations

import os
from typing import Any


def optimization_log_sensitive_enabled() -> bool:
    """Если False (по умолчанию), в print/debug не выводим клиентские/коммерческие поля заказа."""
    v = os.environ.get("OPTIMIZATION_LOG_SENSITIVE", "").strip().lower()
    return v in ("1", "true", "yes", "on")


_REDACT_KEYS = frozenset({
    "customer",
    "kp_id",
    "kp_date",
    "plate_name",
    "client",
    "company",
})


def redact_order(order: dict[str, Any]) -> dict[str, Any]:
    """Копия строки заказа без типичных чувствительных ключей (для логов)."""
    return {k: v for k, v in order.items() if k not in _REDACT_KEYS}


def order_line_for_console(order: dict[str, Any]) -> str:
    """Одна строка для консольного print: либо полная, либо с маскированием."""
    if optimization_log_sensitive_enabled():
        o = order
    else:
        o = redact_order(order)
    qty = o.get("qty", "?")
    length = o.get("length", "?")
    width = o.get("width", "?")
    return f"  {qty}x {length}м x {width}мм"

