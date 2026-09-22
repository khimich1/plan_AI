#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PLI-013: контракт отчёта гейта (orphans / surplus / no_grade)."""
from __future__ import annotations

from core.plan_integrity import build_integrity_report


def _plan_with_item(**item_fields) -> dict:
    item = {
        "length": 6.0,
        "mode": "solid",
        "width": 1.2,
        "load_code": 8,
        "kp_id": 4,
        "plate_name": "ПБ 60-12-8п",
        "kp_plate_id": 10,
        "concrete_grade": "М400",
        **item_fields,
    }
    return {
        "id": "plan_integrity",
        "days": {
            "2026-05-01": {
                "tracks": [{"items": [item]}],
            }
        },
        "optimization_result": {"_coverage_summary": {"surplus": {}}},
    }


def test_clean_plan_has_zero_counts() -> None:
    report = build_integrity_report(_plan_with_item())
    assert report["orphans"] == 0
    assert report["surplus"] == 0
    assert report["no_grade"] == 0
    assert report["items"] == []


def test_item_without_kp_plate_id_counts_as_orphan() -> None:
    report = build_integrity_report(_plan_with_item(kp_plate_id=None))
    assert report["orphans"] == 1
    assert report["items"][0]["kind"] == "orphan"
    assert "kp_plate_id" in report["items"][0]["message"] or "призрак" in report["items"][0]["message"].lower()


def test_item_without_grade_counts_as_no_grade() -> None:
    report = build_integrity_report(_plan_with_item(concrete_grade=""))
    assert report["no_grade"] == 1
    assert report["items"][0]["kind"] == "no_grade"


def test_coverage_surplus_is_counted() -> None:
    plan = _plan_with_item()
    plan["optimization_result"] = {
        "_coverage_summary": {"surplus": {"(6.0, 1200, 8)": 2}},
    }
    report = build_integrity_report(plan)
    assert report["surplus"] == 2
    assert any(entry["kind"] == "surplus" for entry in report["items"])
