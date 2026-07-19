#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Валидация входных данных публичного API оптимизатора (OPT-006)."""

from __future__ import annotations

# Консервативные пределы: не ломаем типичные тесты и прод-порядки.
_MAX_LINES_2D = 2000
_MAX_1D_WIDTH_KEYS = 500
_MAX_PLATE_LENGTH_M = 80.0
_MAX_PLATE_WIDTH_MM = 4000
_MIN_POSITIVE = 1e-6


def validate_optimize_entrypoint(
    *,
    orders: dict | None,
    orders_2d: list | None,
    plate_width: int,
    min_useful_width: int,
) -> None:
    """
    Проверка типов и базовых ограничений до ветвления 1D/2D.

    Пустые orders и orders_2d одновременно — не ошибка (как раньше возвращается {}).
    ``orders=[]`` (список) и прочие неверные типы — ValueError.
    """
    if orders is not None and not isinstance(orders, dict):
        raise ValueError(
            f"orders must be a dict mapping width_mm -> qty, got {type(orders).__name__}. "
            "For 2D mode pass orders_2d=list[dict]."
        )
    if orders_2d is not None and not isinstance(orders_2d, list):
        raise ValueError(
            f"orders_2d must be a list of order dicts, got {type(orders_2d).__name__}."
        )

    if not isinstance(plate_width, int) or plate_width < 200 or plate_width > _MAX_PLATE_WIDTH_MM:
        raise ValueError(f"plate_width must be int in [200, {_MAX_PLATE_WIDTH_MM}], got {plate_width!r}")
    if not isinstance(min_useful_width, int) or min_useful_width < 0 or min_useful_width > plate_width:
        raise ValueError(f"min_useful_width invalid for plate_width={plate_width}: {min_useful_width!r}")

    has_2d = bool(orders_2d) and len(orders_2d) > 0
    has_1d = bool(orders) and len(orders) > 0

    if has_2d:
        if len(orders_2d) > _MAX_LINES_2D:
            raise ValueError(
                f"orders_2d too large: {len(orders_2d)} lines (max {_MAX_LINES_2D}). "
                "Split the job or increase limit in validation module if justified."
            )
        required = ("length", "width", "qty")
        for i, row in enumerate(orders_2d):
            if not isinstance(row, dict):
                raise ValueError(f"orders_2d[{i}] must be dict, got {type(row).__name__}")
            for k in required:
                if k not in row:
                    raise ValueError(f"orders_2d[{i}] missing required key {k!r}")
            L = float(row["length"])
            W = int(row["width"])
            q = row["qty"]
            if L < _MIN_POSITIVE or L > _MAX_PLATE_LENGTH_M:
                raise ValueError(
                    f"orders_2d[{i}].length out of range (0, {_MAX_PLATE_LENGTH_M}]: {L!r}"
                )
            if W < 1 or W > _MAX_PLATE_WIDTH_MM:
                raise ValueError(
                    f"orders_2d[{i}].width out of range [1, {_MAX_PLATE_WIDTH_MM}]: {W!r}"
                )
            try:
                qn = int(q)
            except (TypeError, ValueError) as e:
                raise ValueError(f"orders_2d[{i}].qty must be int-coercible, got {q!r}") from e
            if qn < 1:
                raise ValueError(f"orders_2d[{i}].qty must be >= 1, got {qn}")

    if has_1d:
        if len(orders) > _MAX_1D_WIDTH_KEYS:
            raise ValueError(
                f"orders has too many width keys: {len(orders)} (max {_MAX_1D_WIDTH_KEYS})."
            )
        for w_mm, qty in orders.items():
            try:
                w = int(w_mm)
                q = int(qty)
            except (TypeError, ValueError) as e:
                raise ValueError(f"orders keys must be width_mm (int-like), values qty int-like: {w_mm!r}->{qty!r}") from e
            if w < 1 or w > _MAX_PLATE_WIDTH_MM:
                raise ValueError(f"orders width_mm out of range: {w}")
            if q < 1:
                raise ValueError(f"orders qty must be >= 1 for width {w}: {q}")
