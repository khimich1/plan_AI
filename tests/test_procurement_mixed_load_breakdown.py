"""Группировка и снимок плана для смешанных нагрузок (8п/10п)."""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as _cfg  # noqa: F401

from viz_modules.procurement.breakdown import _accumulate_order_counter
from viz_modules.procurement.plan_snapshot import parse_cascading_plan, snapshot_to_trim_dict
from viz_modules.procurement.trim import _calc_trim_components

BASE_1_2M_69_10 = 24903.0


def test_order_counter_respects_per_row_load_code() -> None:
    orders = [
        {"length": 6.9, "width": 395, "qty": 4, "load_code": 8},
        {"length": 6.9, "width": 395, "qty": 1, "load_code": 10},
    ]
    counter: Counter = Counter()
    _accumulate_order_counter(counter, orders)
    assert counter[(6.9, 395, 8, "", False)] == 4
    assert counter[(6.9, 395, 10, "", False)] == 1
    assert len(counter) == 2


def test_snapshot_preserves_load_code_on_cuts() -> None:
    raw = {
        "primary_cuts": [
            {"width": 395, "rest": 805, "qty": 1, "lengths": [6.9], "load_code": 10},
        ],
        "secondary_cuts": [
            {
                "source": 805,
                "source_lengths": [6.9],
                "lengths": [6.9],
                "cuts": [395],
                "qty": 1,
                "load_code": 8,
            },
        ],
    }
    d = snapshot_to_trim_dict(parse_cascading_plan(raw))
    assert d["primary_cuts"][0]["load_code"] == 10
    assert d["secondary_cuts"][0]["load_code"] == 8


def test_snapshot_preserves_instance_ids_and_waste_410() -> None:
    """Регрессия: ID слэбов не теряются в snapshot; отход 410мм на 10п."""
    raw = {
        "primary_cuts": [
            {
                "width": 395,
                "rest": 805,
                "qty": 1,
                "lengths": [6.9],
                "load_code": 10,
                "primary_instance_id": "prim-2",
            },
        ],
        "secondary_cuts": [
            {
                "source": 805,
                "source_lengths": [6.9],
                "lengths": [6.9],
                "cuts": [395],
                "qty": 1,
                "load_code": 8,
                "parent_instance_id": "prim-2",
            },
            {
                "source": 805,
                "source_lengths": [6.9],
                "lengths": [6.9],
                "cuts": [395],
                "qty": 1,
                "load_code": 8,
                "parent_instance_id": "prim-1",
            },
            {
                "source": 805,
                "source_lengths": [6.9],
                "lengths": [6.9],
                "cuts": [395],
                "qty": 1,
                "load_code": 8,
                "parent_instance_id": "prim-1",
            },
        ],
    }
    plan_dict = snapshot_to_trim_dict(parse_cascading_plan(raw))
    assert plan_dict["primary_cuts"][0]["primary_instance_id"] == "prim-2"
    assert plan_dict["secondary_cuts"][0]["parent_instance_id"] == "prim-2"

    t = _calc_trim_components(
        plan_dict,
        length=6.9,
        width_mm=395,
        qty=1,
        base_price_1_2m=BASE_1_2M_69_10,
        base_price=BASE_1_2M_69_10 * (0.395 / 1.2),
        load_code=10,
        price_table={},
    )
    expected_waste = (410 / 1200.0) * BASE_1_2M_69_10
    assert t["waste_cost"] == pytest.approx(expected_waste)
    assert any(w == 410 for w, _ in t["waste_terms"])
