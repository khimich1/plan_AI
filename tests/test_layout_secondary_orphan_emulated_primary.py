# -*- coding: utf-8 -*-
"""Вторичные без parent_instance_id → эмуляция первичного сплита (вариант A)."""
from __future__ import annotations

import copy

import pytest

from core.config_and_data import make_plate_name
from core.domain.plate_order import normalize_load_code
from core.plate_runtime_state import get_plate_mutable_runtime
from core.production.planning import _build_assignment_gap_fallback_tracks
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.services.optimization_service import OptimizationService
from viz_modules.layout_sequence import _build_sequence_from_plan

from tests.test_layout_secondary_unmatched_parent_user_list import (
    _build_merged_orders_2d,
    _push_plate_load_cfg,
    _restore_plate_load_cfg,
    _norm_lc,
)


@pytest.fixture
def user_list_opt_context():
    orders_2d = _build_merged_orders_2d()
    plate_order = AppPlateOrder.from_orders_2d(orders_2d)
    saved_d, saved_raw = _push_plate_load_cfg(plate_order)
    svc = OptimizationService()
    try:
        ctx = svc.optimize(plate_order, orders_2d=orders_2d)
    finally:
        _restore_plate_load_cfg(saved_d, saved_raw)
    assert ctx.optimization_result
    return orders_2d, plate_order, svc, ctx


def _label(L: float, W: float, load_code: int | None = None) -> str:
    lc = load_code if load_code is not None else 8
    return make_plate_name(L, W, load_code=_norm_lc(lc))


def test_synthetic_orphan_from_minimal_plan_variant_A():
    """Детерминированный план: сирота всегда эмулируется как первичный сплит 1200 − W."""
    plan = {
        "primary_cuts": [
            {
                "width": 1200,
                "rest": 0,
                "qty": 1,
                "lengths": [6.0],
                "load_code": 8,
                "primary_instance_id": "prim-solid-1",
            },
            {
                "width": 530,
                "rest": 670,
                "qty": 1,
                "lengths": [5.98],
                "load_code": 8,
                "primary_instance_id": "prim-split-1",
            },
        ],
        "secondary_cuts": [
            {
                "source": 880,
                "cuts": [320],
                "qty": 1,
                "pieces": 1,
                "waste": 0,
                "type": "narrow",
                "lengths": [5.8],
                "source_lengths": [5.8],
                "target_order_key": (5.8, 720, 8),
                "load_code": 8,
                "parent_instance_id": None,
                "secondary_instance_id": "sec-orphan-1",
                "kp_id": 101,
                "plate_name": "ПБ тест сирота",
            },
        ],
        "plate_assignments": [
            {
                "unit_id": "sec-orphan-1",
                "source": "secondary",
                "length": 5.8,
                "width": 320,
                "kp_id": 101,
                "plate_name": "ПБ тест сирота",
                "load_code": 8,
            },
        ],
    }

    flat = _build_sequence_from_plan(plan, _label, reinforcement_map={})
    hit = next(
        (x for x in flat if x.get("unit_id") == "sec-orphan-1" and x.get("mode") == "split"),
        None,
    )
    assert hit is not None
    assert not (hit.get("secondary_cuts") or [])
    assert abs(hit["main_w"] - 0.32) < 1e-6
    assert abs(hit["rest_w"] - (1200 - 320) / 1000.0) < 1e-6

    tracks = [{"label": "Д1", "items": flat, "length": 100.0}]
    missing_fb, _ = _build_assignment_gap_fallback_tracks(
        plate_assignments=plan["plate_assignments"],
        tracks_list=tracks,
    )
    assert missing_fb == []


def test_orphan_secondary_emulated_split_has_sec_unit_id_and_rest_1200_minus_width(
    user_list_opt_context,
):
    _o, plate_order, svc, ctx = user_list_opt_context
    res = copy.deepcopy(ctx.optimization_result or {})
    orphans = [c for c in (res.get("secondary_cuts") or []) if not c.get("parent_instance_id")]
    if not orphans:
        pytest.skip("нет secondary без parent — эмуляция не проверяется")
    orphan = orphans[0]
    sid = orphan.get("secondary_instance_id")
    assert sid is not None
    out_w = int((orphan.get("cuts") or [0])[0])
    assert 0 < out_w < 1200
    expected_rest_mm = 1200 - out_w

    saved_d, saved_raw = _push_plate_load_cfg(plate_order)
    try:
        flat = _build_sequence_from_plan(res, _label, reinforcement_map={})
    finally:
        _restore_plate_load_cfg(saved_d, saved_raw)

    matches = [
        x
        for x in flat
        if x.get("mode") == "split"
        and str(x.get("unit_id")) == str(sid)
        and not (x.get("secondary_cuts") or [])
    ]
    assert matches, "ожидали split без secondary_cuts с unit_id=secondary_instance_id сироты"
    item = matches[0]
    assert abs(float(item["main_w"]) - out_w / 1000.0) < 1e-6
    assert abs(float(item["rest_w"]) - expected_rest_mm / 1000.0) < 1e-6


def test_orphan_emulated_split_not_in_fallback_track(user_list_opt_context):
    _o, plate_order, svc, ctx = user_list_opt_context
    res = copy.deepcopy(ctx.optimization_result or {})
    orphans = [c for c in (res.get("secondary_cuts") or []) if not c.get("parent_instance_id")]
    if not orphans:
        pytest.skip("нет secondary без parent")
    sid = str(orphans[0].get("secondary_instance_id"))

    saved_d, saved_raw = _push_plate_load_cfg(plate_order)
    try:
        seq = _build_sequence_from_plan(res, _label, reinforcement_map={})
        tracks = [
            {
                "label": "Д1",
                "items": seq,
                "length": 100.0,
            }
        ]
        assignments = list(res.get("plate_assignments") or [])
        missing_fb, _ctr = _build_assignment_gap_fallback_tracks(
            plate_assignments=assignments,
            tracks_list=tracks,
        )
        sec_assign_ids = {
            str(a.get("unit_id"))
            for a in assignments
            if a.get("source") == "secondary" and a.get("unit_id")
        }
        if sid not in sec_assign_ids:
            pytest.skip("sec id не в plate_assignments")
        assert missing_fb == [], (
            f"сирота {sid} должна быть в основной дорожке, не FALLBACK: got {missing_fb!r}"
        )
    finally:
        _restore_plate_load_cfg(saved_d, saved_raw)
