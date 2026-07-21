# -*- coding: utf-8 -*-
"""Жадная перестановка ordered_cuts и сплиттер с предпочтением армирования."""

from __future__ import annotations

import pytest
from core.optimization.layout_runtime_snapshot import LayoutSequenceCfgSlice
from core.plate_runtime_state import get_plate_mutable_runtime
from core.visualization import split_sequence_into_tracks
from viz_modules.layout_sequence.from_plan import _build_sequence_from_plan


def _label(L: float, W: float, load_code: int | None = None) -> str:
    return f"{L:.2f}×{W}"


def _layout_cfg(*, greedy: bool, track_pref: bool = False) -> LayoutSequenceCfgSlice:
    return LayoutSequenceCfgSlice.from_plate_runtime(
        get_plate_mutable_runtime(),
        layout_greedy_reinf_merge=greedy,
        layout_track_reinf_preference=track_pref,
    )


@pytest.fixture
def plan_two_solids_two_cut_groups() -> dict:
    """Две целые (разное армирование) и две группы с резом; целые легче первой группы резов."""
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
def reinforcement_map_plan_two() -> dict:
    return {
        (6.0, 1200, 8): 1.0,
        (5.0, 1200, 8): 3.0,
        (6.0, 1200, 9): 10.0,
        (6.0, 1200, 12): 50.0,
    }


def test_greedy_inserts_light_solid_before_splits(
    plan_two_solids_two_cut_groups: dict,
    reinforcement_map_plan_two: dict,
) -> None:
    legacy = _build_sequence_from_plan(
        plan_two_solids_two_cut_groups,
        _label,
        reinforcement_map_plan_two,
        layout_cfg=_layout_cfg(greedy=False),
    )
    greedy = _build_sequence_from_plan(
        plan_two_solids_two_cut_groups,
        _label,
        reinforcement_map_plan_two,
        layout_cfg=_layout_cfg(greedy=True),
    )
    assert legacy[0]["mode"] == "solid"
    assert greedy[0]["mode"] == "solid"
    # Legacy: вторая запись — уже сплит первой группы
    assert legacy[1]["mode"] == "split"
    # Greedy: вторая — вторая целая (5 м), сплиты отложены
    assert greedy[1]["mode"] == "solid"
    assert greedy[2]["mode"] == "split"


def test_greedy_then_split_track_item_count_integrity(
    plan_two_solids_two_cut_groups: dict,
    reinforcement_map_plan_two: dict,
) -> None:
    seq = _build_sequence_from_plan(
        plan_two_solids_two_cut_groups,
        _label,
        reinforcement_map_plan_two,
        layout_cfg=_layout_cfg(greedy=True),
    )
    tracks = split_sequence_into_tracks(
        seq,
        strict_layout_integrity=True,
        track_reinf_preference=True,
        track_start_reinf_relaxation=True,
    )
    assert sum(len(t["items"]) for t in tracks) == len(seq)
    for tr in tracks:
        if tr.get("items"):
            assert tr["items"][0].get("mode") == "solid"


def test_track_reinforcement_spread_metric_helpers() -> None:
    """Документация приёмки: разброс армирования по дорожкам (max−min по items)."""
    seq = [
        {"length": 10.0, "mode": "solid", "reinforcement": 5.0},
        {"length": 10.0, "mode": "split", "main_w": 1.2, "rest_w": 0.0, "reinforcement": 40.0},
    ]
    tracks = split_sequence_into_tracks(seq, track_reinf_preference=False)
    reinfs_per_track = []
    for tr in tracks:
        vals = [float(it.get("reinforcement") or 0) for it in tr["items"]]
        if vals:
            reinfs_per_track.append(max(vals) - min(vals))
    assert reinfs_per_track
    assert reinfs_per_track[0] >= 0
