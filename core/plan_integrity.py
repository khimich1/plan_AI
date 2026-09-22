#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Отчёт целостности плана для plan-view (orphans / surplus / no_grade)."""

from __future__ import annotations

from typing import Any, Iterable

INTEGRITY_KEY = "integrity"


def empty_integrity_report() -> dict[str, Any]:
    return {"orphans": 0, "surplus": 0, "no_grade": 0, "items": []}


def _iter_physical_items(plan: dict[str, Any]) -> Iterable[dict[str, Any]]:
    days = plan.get("days") or {}
    if not isinstance(days, dict):
        return
    for day in days.values():
        if not isinstance(day, dict):
            continue
        for track in day.get("tracks") or []:
            if not isinstance(track, dict):
                continue
            for item in track.get("items") or []:
                if not isinstance(item, dict):
                    continue
                yield item
                for sec in item.get("secondary_cuts") or []:
                    if isinstance(sec, dict):
                        yield sec


def _has_identity(item: dict[str, Any]) -> bool:
    return bool(item.get("kp_id") and (item.get("plate_name") or item.get("label")))


def _grade_missing(item: dict[str, Any]) -> bool:
    raw = item.get("concrete_grade")
    if raw is None:
        return True
    text = str(raw).strip()
    return text == "" or text == "—"


def _surplus_total(plan: dict[str, Any]) -> tuple[int, list[dict[str, str]]]:
    opt = plan.get("optimization_result") or {}
    if not isinstance(opt, dict):
        return 0, []
    cov = opt.get("_coverage_summary") or {}
    if not isinstance(cov, dict):
        return 0, []
    surplus = cov.get("surplus") or {}
    if not isinstance(surplus, dict):
        return 0, []
    total = 0
    items: list[dict[str, str]] = []
    for key, qty in surplus.items():
        try:
            count = int(qty)
        except (TypeError, ValueError):
            continue
        if count <= 0:
            continue
        total += count
        items.append(
            {
                "kind": "surplus",
                "message": f"Лишние плиты: {key} ×{count}",
            }
        )
    return total, items


def build_integrity_report(plan: dict[str, Any] | None) -> dict[str, Any]:
    """Считает orphans / surplus / no_grade по JSON плана (последний снимок)."""
    report = empty_integrity_report()
    if not plan:
        return report

    items: list[dict[str, str]] = []
    orphans = 0
    no_grade = 0
    for physical in _iter_physical_items(plan):
        if not _has_identity(physical):
            continue
        name = physical.get("plate_name") or physical.get("label") or ""
        kp = physical.get("kp_id")
        if not physical.get("kp_plate_id"):
            orphans += 1
            items.append(
                {
                    "kind": "orphan",
                    "message": f"призрак: нет kp_plate_id (КП #{kp}, {name})",
                }
            )
        if _grade_missing(physical):
            no_grade += 1
            items.append(
                {
                    "kind": "no_grade",
                    "message": f"Нет марки: КП #{kp}, {name}",
                }
            )

    surplus, surplus_items = _surplus_total(plan)
    items.extend(surplus_items)
    report["orphans"] = orphans
    report["surplus"] = surplus
    report["no_grade"] = no_grade
    report["items"] = items
    return report
