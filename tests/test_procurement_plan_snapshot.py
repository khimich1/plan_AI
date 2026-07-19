from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as _cfg  # noqa: F401 — стабилизирует порядок загрузки core ↔ viz_modules

from viz_modules.procurement.orders import get_orders_from_opt_plan
from viz_modules.procurement.plan_snapshot import (
    CascadingPlanSnapshot,
    PlanSnapshotValidationError,
    parse_cascading_plan,
    snapshot_to_trim_dict,
)
from core.optimization.layout_runtime_snapshot import OptPlanFrozenSnapshot


def test_parse_empty_plan() -> None:
    assert parse_cascading_plan(None) == CascadingPlanSnapshot()
    assert parse_cascading_plan({}).orders_requested == []


def test_parse_orders_coercion() -> None:
    snap = parse_cascading_plan(
        {
            "orders_requested": [
                {"length": "7.3", "width": 320, "qty": "2", "load_code": 8, "length_dm_raw": " 73 "},
            ],
        }
    )
    assert len(snap.orders_requested) == 1
    row = snap.orders_requested[0]
    assert row.length == pytest.approx(7.3)
    assert row.width == 320
    assert row.qty == 2
    assert row.length_dm_raw == "73"


def test_invalid_orders_requested_type_raises() -> None:
    with pytest.raises(PlanSnapshotValidationError):
        parse_cascading_plan({"orders_requested": "broken"})


def test_snapshot_trim_dict_matches_trim_contract() -> None:
    snap = parse_cascading_plan(
        {
            "primary_cuts": [{"width": 1200, "lengths": [5.71], "qty": 1, "rest": 100}],
            "secondary_cuts": [
                {
                    "source": 100,
                    "source_lengths": [5.71],
                    "lengths": [5.71],
                    "cuts": [900],
                    "qty": 1,
                    "pieces": 1,
                    "waste": 12.5,
                    "type": "transverse",
                }
            ],
        }
    )
    d = snapshot_to_trim_dict(snap)
    assert d["primary_cuts"][0]["width"] == 1200
    assert d["primary_cuts"][0]["lengths"][0] == pytest.approx(5.71)
    assert d["secondary_cuts"][0]["cuts"][0] == 900


def test_get_orders_from_opt_snapshot_no_globals() -> None:
    frozen = OptPlanFrozenSnapshot(
        opt_plan={},
        opt_cascading_plan={
            "orders_requested": [
                {"length": 6.0, "width": 1200, "qty": 3},
            ]
        },
        opt_cascading_plan_by_load={},
        opt_width_priority=(),
        load_to_reinforcement_map={},
    )
    orders = get_orders_from_opt_plan(opt_snapshot=frozen)
    assert orders is not None
    assert len(orders) == 1
    assert orders[0]["length"] == pytest.approx(6.0)
    assert orders[0]["qty"] == 3


def test_by_load_priority_over_single_plan() -> None:
    frozen = OptPlanFrozenSnapshot(
        opt_plan={},
        opt_cascading_plan={
            "orders_requested": [{"length": 99.0, "width": 1200, "qty": 1}],
        },
        opt_cascading_plan_by_load={
            8: {"orders_requested": [{"length": 5.0, "width": 1150, "qty": 2}]},
        },
        opt_width_priority=(),
        load_to_reinforcement_map={},
    )
    orders = get_orders_from_opt_plan(opt_snapshot=frozen)
    assert orders is not None
    assert orders[0]["length"] == pytest.approx(5.0)
    assert orders[0]["width"] == 1150
