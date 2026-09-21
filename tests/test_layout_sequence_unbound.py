"""Regression tests for 2D layout length handling."""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.optimization as optimization  # noqa: E402
from viz_modules.layout_sequence import (  # noqa: E402
    _build_sequence_from_plan,
    build_layout_sequence,
)


def _minimal_2d_plan() -> dict:
    return {
        "plate_assignments": [
            {
                "plate_uid": "plate-1",
                "kp_id": 1,
                "plate_name": "Плита ПБ 55-12-8п",
                "length": 5.5,
                "width": 1200,
                "load_code": 8,
            }
        ],
        "primary_cuts": [
            {
                "qty": 1,
                "width": 1200,
                "rest": 0,
                "lengths": [5.5],
                "plate_uids": ["plate-1"],
                "load_code": 8,
                "kp_id": 1,
                "plate_name": "Плита ПБ 55-12-8п",
            }
        ],
        "secondary_cuts": [],
        "transverse_cuts": [],
    }


def _plate_label(length: float, width_m: float, load_code: int | None = None) -> str:
    return f"{length:g}x{int(round(width_m * 1000))}-{load_code or 8}"


def test_build_sequence_from_plan_defines_legacy_lengths_in_2d_mode() -> None:
    sequence = _build_sequence_from_plan(_minimal_2d_plan(), _plate_label, {})

    assert len(sequence) == 1
    assert sequence[0]["mode"] == "solid"
    assert sequence[0]["length"] == 5.5
    assert sequence[0]["plate_uid"] == "plate-1"


def test_build_layout_sequence_defines_legacy_lengths_in_2d_mode(monkeypatch) -> None:
    monkeypatch.setattr(optimization, "OPT_CASCADING_PLAN_BY_LOAD", None)
    monkeypatch.setattr(optimization, "OPT_CASCADING_PLAN", _minimal_2d_plan())

    sequence = build_layout_sequence()

    assert len(sequence) == 1
    assert sequence[0]["mode"] == "solid"
    assert sequence[0]["length"] == 5.5
    assert sequence[0]["plate_uid"] == "plate-1"


def test_build_sequence_copies_concrete_grade_on_solid_item() -> None:
    plan = _minimal_2d_plan()
    plan["primary_cuts"][0]["concrete_grade"] = "М500"

    sequence = _build_sequence_from_plan(plan, _plate_label, {})

    assert sequence[0]["mode"] == "solid"
    assert sequence[0]["concrete_grade"] == "М500"


def test_build_sequence_copies_concrete_grade_on_split_item() -> None:
    plan = {
        "plate_assignments": [True],
        "primary_cuts": [
            {
                "qty": 1,
                "width": 1200,
                "rest": 200,
                "lengths": [6.0],
                "plate_uids": ["plate-split"],
                "load_code": 8,
                "kp_id": 2,
                "plate_name": "Плита ПБ 60-12-8п",
                "concrete_grade": "М400",
            }
        ],
        "secondary_cuts": [],
        "transverse_cuts": [],
    }

    sequence = _build_sequence_from_plan(plan, _plate_label, {})

    assert sequence[0]["mode"] == "split"
    assert sequence[0]["concrete_grade"] == "М400"


def test_build_sequence_copies_concrete_grade_on_transverse_item() -> None:
    plan = {
        "plate_assignments": [True],
        "primary_cuts": [
            {
                "qty": 1,
                "width": 1200,
                "rest": 0,
                "lengths": [4.2],
                "plate_uids": ["plate-transverse"],
                "load_code": 8,
                "kp_id": 3,
                "plate_name": "Плита ПБ 42-12-8п",
                "concrete_grade": "М500",
            }
        ],
        "secondary_cuts": [],
        "transverse_cuts": [
            {
                "source_length": 4.2,
                "source_width": 1200,
                "target_length": 3.0,
                "remainder": 1.2,
            }
        ],
    }

    sequence = _build_sequence_from_plan(plan, _plate_label, {})

    assert sequence[0]["mode"] == "transverse"
    assert sequence[0]["concrete_grade"] == "М500"
