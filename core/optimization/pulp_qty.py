#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Извлечение неотрицательного целого qty из решённых PuLP-переменных (ветка 1D)."""

from __future__ import annotations

import logging
from typing import Any, Callable

_LOGGER = logging.getLogger(__name__)


def _opt_1d_pulp_nonneg_qty(
    pulp_value_fn: Callable[[Any], Any],
    var: Any,
    *,
    context: str,
) -> int:
    """
    Целое qty ≥ 0 из решённой PuLP-переменной (ветка 1D).

    ``None`` от value() → 0 и предупреждение (нестабильное/частичное решение без молчаливой порчи списка).
    Ошибки преобразования и сбои value() логируются и превращаются в ValueError.
    """
    try:
        raw = pulp_value_fn(var)
    except Exception as exc:
        _LOGGER.exception(
            "[OPT_1D] pulp.value() выбросил исключение для %s",
            context,
        )
        raise ValueError(f"pulp.value failed for {context}") from exc
    if raw is None:
        _LOGGER.warning(
            "[OPT_1D] %s: value() вернул None после решения — qty=0",
            context,
        )
        return 0
    try:
        qty = int(round(float(raw)))
    except (TypeError, ValueError, OverflowError) as exc:
        _LOGGER.error(
            "[OPT_1D] %s: не удалось преобразовать value=%r в int",
            context,
            raw,
            exc_info=True,
        )
        raise ValueError(f"invalid pulp value for {context}: {raw!r}") from exc
    if qty < 0:
        _LOGGER.error("[OPT_1D] %s: отрицательное qty=%s", context, qty)
        raise ValueError(f"negative qty for {context}: {qty}")
    return qty
