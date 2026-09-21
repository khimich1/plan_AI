"""Hash раскладки по layout-полям (D-hash, PLI-001).

Hash игнорирует атрибутивные поля ``concrete_grade`` / ``kp_id`` / ``kp_plate_id``,
чтобы propagation марки не ломал приёмку.
"""
from __future__ import annotations

import copy
from pathlib import Path

from core.layout_hash import (
    ATTRIBUTION_KEYS,
    golden_plan,
    hash_layout_sequence,
    layout_canonical,
)
from viz_modules.layout_sequence.from_plan import _build_sequence_from_plan

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _label(length: float, width_m: float, load_code: int | None = None) -> str:
    return f"{length:g}x{int(round(width_m * 1000))}-{load_code or 8}"


def _build() -> list[dict]:
    return _build_sequence_from_plan(golden_plan(), _label, {})


def test_two_runs_on_golden_plan_give_same_hash() -> None:
    first = hash_layout_sequence(_build())
    second = hash_layout_sequence(_build())
    assert first == second
    assert len(first) == 64
    baseline_path = PROJECT_ROOT / "scripts" / "layout_hash_baseline.sha256"
    stored = baseline_path.read_text(encoding="utf-8").strip()
    assert first == stored


def test_adding_attribution_fields_does_not_change_hash() -> None:
    sequence = _build()
    baseline = hash_layout_sequence(sequence)

    mutated = copy.deepcopy(sequence)
    for item in mutated:
        item["concrete_grade"] = "М500"
        item["kp_id"] = 999
        item["kp_plate_id"] = 42
        for sec in item.get("secondary_cuts") or []:
            if isinstance(sec, dict):
                sec["concrete_grade"] = "М400"
                sec["kp_id"] = 888
                sec["kp_plate_id"] = 7

    assert hash_layout_sequence(mutated) == baseline


def test_layout_canonical_drops_attribution_keys() -> None:
    item = {
        "length": 6.0,
        "mode": "solid",
        "width": 1.2,
        "load_code": 8,
        "concrete_grade": "М500",
        "kp_id": 4,
        "kp_plate_id": 197,
        "plate_name": "ПБ 60-12-8п",
    }
    canonical = layout_canonical([item])
    dumped = canonical[0]
    for key in ATTRIBUTION_KEYS:
        assert key not in dumped
    assert dumped["length"] == 6.0
    assert dumped["mode"] == "solid"
    assert dumped["load_code"] == 8
    assert dumped["plate_name"] == "ПБ 60-12-8п"


def test_geometry_change_changes_hash() -> None:
    sequence = _build()
    baseline = hash_layout_sequence(sequence)
    mutated = copy.deepcopy(sequence)
    mutated[0]["length"] = float(mutated[0]["length"]) + 0.1
    assert hash_layout_sequence(mutated) != baseline


def test_golden_root_items_carry_cut_grade_not_resolver_default() -> None:
    plan = golden_plan()
    expected = {
        cut["plate_name"]: cut["concrete_grade"]
        for cut in plan["primary_cuts"]
    }
    sequence = _build()
    assert sequence
    for item in sequence:
        name = item.get("plate_name")
        grade = item.get("concrete_grade")
        assert grade, f"нет марки у item {name!r}"
        if name in expected:
            assert grade == expected[name]
    # ПБ 42 resolver дефолтит в М400; в эталоне заказ — М500.
    transverse = next(item for item in sequence if item.get("mode") == "transverse")
    assert transverse["concrete_grade"] == "М500"
