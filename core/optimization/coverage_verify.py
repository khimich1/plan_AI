#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Проверка покрытия спроса по primary_cuts + secondary_cuts."""

from __future__ import annotations

from collections import Counter

from core.config_and_data import canonical_plate_key


def verify_coverage(
    demand_2d: dict,
    primary_cuts: list,
    secondary_cuts: list,
) -> dict:
    """
    Сверяет фактическое покрытие спроса по primary_cuts + secondary_cuts.

    Returns:
        {
            "demand_total": int,
            "covered_total": int,
            "missing": {(L, W, lc): int},     # дефицит по ключам
            "surplus": {(L, W, lc): int},     # перепроизводство по ключам
            "ok": bool                          # True, если нет дефицита
        }
    Все ключи приведены через canonical_plate_key, так что 800/8 и
    дробные длины не дают ложного несоответствия.
    """
    demand_norm: Counter = Counter()
    for key, qty in (demand_2d or {}).items():
        if isinstance(key, tuple) and len(key) == 3:
            demand_norm[canonical_plate_key(*key)] += int(qty)

    coverage: Counter = Counter()
    for cut in primary_cuts or []:
        ak = cut.get("assignment_key")
        if ak and isinstance(ak, tuple) and len(ak) == 3:
            coverage[canonical_plate_key(*ak)] += 1
    for cut in secondary_cuts or []:
        tk = cut.get("target_order_key")
        if tk and isinstance(tk, tuple) and len(tk) == 3:
            coverage[canonical_plate_key(*tk)] += 1

    missing: dict = {}
    surplus: dict = {}
    for key, need in demand_norm.items():
        have = coverage.get(key, 0)
        if have < need:
            missing[key] = need - have
        elif have > need:
            surplus[key] = have - need

    return {
        "demand_total": int(sum(demand_norm.values())),
        "covered_total": int(sum(coverage.values())),
        "missing": missing,
        "surplus": surplus,
        "ok": not missing,
    }
