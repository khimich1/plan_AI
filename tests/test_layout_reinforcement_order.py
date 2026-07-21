# -*- coding: utf-8 -*-
"""Режим layout_reinforcement_order: asc vs desc (max-first + match-greedy)."""

from __future__ import annotations

import pytest
from core.optimization.layout_runtime_snapshot import (
    LayoutSequenceCfgSlice,
    build_layout_runtime_snapshot,
)
from core.plate_runtime_state import get_plate_mutable_runtime
from viz_modules.layout_sequence.from_plan import _build_sequence_from_plan
from viz_modules.layout_sequence.helpers import choose_best_separator, choose_closest_solid


def _label(L: float, W: float, load_code: int | None = None) -> str:
    return f"{L:.2f}×{W}"


def _layout_cfg(
    *,
    reinforcement_order: str = "asc",
    greedy: bool = True,
) -> LayoutSequenceCfgSlice:
    return LayoutSequenceCfgSlice.from_plate_runtime(
        get_plate_mutable_runtime(),
        layout_greedy_reinf_merge=greedy,
        layout_reinforcement_order=reinforcement_order,  # type: ignore[arg-type]
    )


@pytest.fixture
def plan_mixed_solids_and_cuts() -> dict:
    """Две целые (1.0 и 3.0 кг/м) и две группы с резом (10 и 50 кг/м)."""
    return {
        "plate_assignments": [True],
        "primary_cuts": [
            {"width": 1200, "rest": 0, "qty": 1, "lengths": [6.0], "load_code": 8},
            {"width": 1200, "rest": 0, "qty": 1, "lengths": [5.0], "load_code": 8},
            {"width": 1200, "rest": 200, "qty": 1, "lengths": [6.0], "load_code": 9},
            {"width": 1200, "rest": 300, "qty": 1, "lengths": [6.0], "load_code": 12},
        ],
        "secondary_cuts": [],
    }


@pytest.fixture
def reinforcement_map_mixed() -> dict:
    return {
        (6.0, 1200, 8): 1.0,
        (5.0, 1200, 8): 3.0,
        (6.0, 1200, 9): 10.0,
        (6.0, 1200, 12): 50.0,
    }


@pytest.fixture
def plan_tiered_solids_and_cuts() -> dict:
    """Целые 6/24/36 и группы резов 6 и 24 — для проверки match-greedy."""
    return {
        "plate_assignments": [True],
        "primary_cuts": [
            {"width": 1200, "rest": 0, "qty": 1, "lengths": [6.0], "load_code": 8},
            {"width": 1200, "rest": 0, "qty": 1, "lengths": [6.0], "load_code": 9},
            {"width": 1200, "rest": 0, "qty": 1, "lengths": [6.0], "load_code": 12},
            {"width": 1200, "rest": 200, "qty": 1, "lengths": [6.0], "load_code": 8},
            {"width": 1200, "rest": 200, "qty": 1, "lengths": [6.0], "load_code": 9},
        ],
        "secondary_cuts": [],
    }


@pytest.fixture
def reinforcement_map_tiered() -> dict:
    return {
        (6.0, 1200, 8): 6.0,
        (6.0, 1200, 9): 24.0,
        (6.0, 1200, 12): 36.0,
    }


def _solid_reinforcements(seq: list[dict]) -> list[float]:
    return [float(s.get("reinforcement") or 0) for s in seq if s.get("mode") == "solid"]


def _neighbor_solid_before_split(seq: list[dict], split_reinf: float) -> float | None:
    for i, item in enumerate(seq):
        if item.get("mode") != "split":
            continue
        if abs(float(item.get("reinforcement") or 0) - split_reinf) > 0.5:
            continue
        for j in range(i - 1, -1, -1):
            if seq[j].get("mode") == "solid":
                return float(seq[j].get("reinforcement") or 0)
        return None
    return None


def test_desc_first_solid_is_max_reinforcement(
    plan_mixed_solids_and_cuts: dict,
    reinforcement_map_mixed: dict,
) -> None:
    seq = _build_sequence_from_plan(
        plan_mixed_solids_and_cuts,
        _label,
        reinforcement_map_mixed,
        layout_cfg=_layout_cfg(reinforcement_order="desc", greedy=False),
    )
    solids = _solid_reinforcements(seq)
    assert solids, "ожидаем хотя бы одну целую в sequence"
    assert solids[0] == max(solids)


def test_desc_trailing_solids_are_weaker_than_first(
    plan_mixed_solids_and_cuts: dict,
    reinforcement_map_mixed: dict,
) -> None:
    seq = _build_sequence_from_plan(
        plan_mixed_solids_and_cuts,
        _label,
        reinforcement_map_mixed,
        layout_cfg=_layout_cfg(reinforcement_order="desc", greedy=False),
    )
    solids = _solid_reinforcements(seq)
    if len(solids) >= 2:
        assert min(solids[1:]) <= solids[0]
        assert min(solids) == 1.0


def test_desc_match_places_closest_solid_before_split_group(
    plan_tiered_solids_and_cuts: dict,
    reinforcement_map_tiered: dict,
) -> None:
    seq = _build_sequence_from_plan(
        plan_tiered_solids_and_cuts,
        _label,
        reinforcement_map_tiered,
        layout_cfg=_layout_cfg(reinforcement_order="desc", greedy=False),
    )
    neighbor = _neighbor_solid_before_split(seq, split_reinf=6.0)
    assert neighbor is not None
    assert abs(neighbor - 6.0) < abs(neighbor - 36.0)


def test_desc_asc_greedy_cfg_does_not_change_desc_match_order(
    plan_mixed_solids_and_cuts: dict,
    reinforcement_map_mixed: dict,
) -> None:
    """desc использует match-greedy; asc-greedy флаг в cfg на desc не влияет."""
    with_greedy_flag = _build_sequence_from_plan(
        plan_mixed_solids_and_cuts,
        _label,
        reinforcement_map_mixed,
        layout_cfg=_layout_cfg(reinforcement_order="desc", greedy=True),
    )
    without_greedy_flag = _build_sequence_from_plan(
        plan_mixed_solids_and_cuts,
        _label,
        reinforcement_map_mixed,
        layout_cfg=_layout_cfg(reinforcement_order="desc", greedy=False),
    )
    assert [s.get("mode") for s in with_greedy_flag] == [
        s.get("mode") for s in without_greedy_flag
    ]


def test_desc_match_differs_from_asc_linear(
    plan_tiered_solids_and_cuts: dict,
    reinforcement_map_tiered: dict,
) -> None:
    desc_seq = _build_sequence_from_plan(
        plan_tiered_solids_and_cuts,
        _label,
        reinforcement_map_tiered,
        layout_cfg=_layout_cfg(reinforcement_order="desc", greedy=False),
    )
    asc_linear = _build_sequence_from_plan(
        plan_tiered_solids_and_cuts,
        _label,
        reinforcement_map_tiered,
        layout_cfg=_layout_cfg(reinforcement_order="asc", greedy=False),
    )
    assert [s.get("mode") for s in desc_seq] != [s.get("mode") for s in asc_linear]


def test_build_layout_runtime_snapshot_desc_forces_greedy_off() -> None:
    rt = build_layout_runtime_snapshot(layout_reinforcement_order="desc")
    assert rt.layout_cfg.layout_reinforcement_order == "desc"
    assert rt.layout_cfg.layout_greedy_reinf_merge is False


def test_choose_closest_solid_picks_min_distance() -> None:
    solid_pool = [
        {"lengths": [6.0], "width": 1200},
        {"lengths": [5.0], "width": 1200},
        {"lengths": [4.0], "width": 1200},
    ]
    reinforcement_map = {
        (6.0, 1200, 8): 5.0,
        (5.0, 1200, 8): 20.0,
        (4.0, 1200, 8): 8.0,
    }
    idx = choose_closest_solid(solid_pool, 6.0, reinforcement_map)
    assert idx == 0


def test_choose_best_separator_asc_vs_desc() -> None:
    solid_pool = [
        {"lengths": [6.0], "width": 1200, "reinforcement": 5.0},
        {"lengths": [5.0], "width": 1200, "reinforcement": 20.0},
        {"lengths": [4.0], "width": 1200, "reinforcement": 8.0},
    ]
    next_group = [{"reinforcement": 6.0}]
    reinforcement_map = {
        (6.0, 1200, 8): 5.0,
        (5.0, 1200, 8): 20.0,
        (4.0, 1200, 8): 8.0,
    }
    asc_idx = choose_best_separator(
        solid_pool, next_group, reinforcement_map, reinforcement_order="asc"
    )
    desc_idx = choose_best_separator(
        solid_pool, next_group, reinforcement_map, reinforcement_order="desc"
    )
    assert asc_idx == 0
    assert desc_idx == 0
    assert solid_pool[asc_idx]["reinforcement"] == 5.0
    assert solid_pool[desc_idx]["reinforcement"] == 5.0


def test_desc_separator_closest_not_max() -> None:
    solid_pool = [
        {"lengths": [6.0], "width": 1200},
        {"lengths": [5.0], "width": 1200},
        {"lengths": [4.0], "width": 1200},
    ]
    next_group = [{"reinforcement": 6.0}]
    reinforcement_map = {
        (6.0, 1200, 8): 5.0,
        (5.0, 1200, 8): 20.0,
        (4.0, 1200, 8): 8.0,
    }
    desc_idx = choose_best_separator(
        solid_pool, next_group, reinforcement_map, reinforcement_order="desc"
    )
    assert desc_idx == 0
    assert reinforcement_map[(6.0, 1200, 8)] == 5.0
