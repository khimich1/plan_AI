# -*- coding: utf-8 -*-
from __future__ import annotations

import copy
import logging

import pytest

import core.config_and_data as cfg
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.services.optimization_service import OptimizationService
from core.optimization import verify_coverage
from core.optimization.secondary_batches import _batch_sizes_for_secondary_z_sec
from core.plate_line_parser import parse_line

LINES = """
Плиты ПБ 23-6,65-8п  3
Плиты ПБ 51-5,3-8п  1
Плиты ПБ 30-8,0-8п  1
Плиты ПБ 53-5,3-8п  1
Плиты ПБ 88-3,2-8п  2
Плиты ПБ 23-6,65-8п  1
Плиты ПБ 28-7,2-8п  1
Плиты ПБ 50-6,65-8п  2
Плиты ПБ 58-9,2-8п  1
Плиты ПБ 59-12-8п  2
Плиты ПБ 90-6,65-8п  1
Плиты ПБ 59-7,2-8п  2
Плиты ПБ 58-3,2-8п  3
Плиты ПБ 58-3,2-8п  1
Плиты ПБ 58-7,2-8п  1
Плиты ПБ 52-3,2-8п  1
Плиты ПБ 80-12-10п  5
Плиты ПБ 80-9,2-10п  1
Плиты ПБ 80-3,2-10п  1
""".strip().splitlines()


def _norm_lc(value: float | int | None) -> int:
    return cfg.normalize_load_code(value, default=8)


def _build_merged_orders_2d() -> list[dict]:
    rows: list[dict] = []
    for raw in LINES:
        line = raw.strip()
        if not line:
            continue
        parsed = parse_line(line)
        assert parsed.parsed, (line, parsed.reason_text)
        lc = parsed.load_code if parsed.load_code is not None else 8.0
        rows.append(
            {
                "length": float(parsed.length_m),
                "width": int(round(parsed.width_m * 1000)),
                "qty": int(parsed.qty),
                "load_code": float(lc),
                "length_dm_raw": parsed.length_dm_raw or "",
            }
        )

    merged: dict[tuple[float, int, int], int] = {}
    for item in rows:
        key = (round(item["length"], 4), item["width"], _norm_lc(item["load_code"]))
        merged[key] = merged.get(key, 0) + item["qty"]

    return [
        {"length": key[0], "width": key[1], "qty": qty, "load_code": float(key[2])}
        for key, qty in sorted(merged.items())
    ]


def _push_plate_load_cfg(plate_order: AppPlateOrder) -> tuple[dict, dict]:
    saved_details = copy.deepcopy(cfg.PLATE_LOAD_DETAILS)
    saved_dm_raw = copy.deepcopy(cfg.PLATE_LENGTH_DM_RAW)

    cfg.PLATE_LOAD_DETAILS.clear()
    for key, qty in plate_order.plate_load_details.items():
        length, width_m, load_code, raw = key
        cfg.PLATE_LOAD_DETAILS[(length, width_m, int(float(load_code)), raw)] = int(qty)

    cfg.PLATE_LENGTH_DM_RAW.clear()
    for key, raw in plate_order.plate_length_dm_raw.items():
        length, width_m, load_code, raw_val = key
        cfg.PLATE_LENGTH_DM_RAW[(length, width_m, int(float(load_code)), raw_val)] = raw

    return saved_details, saved_dm_raw


def _restore_plate_load_cfg(saved_details: dict, saved_dm_raw: dict) -> None:
    cfg.PLATE_LOAD_DETAILS.clear()
    cfg.PLATE_LOAD_DETAILS.update(saved_details)
    cfg.PLATE_LENGTH_DM_RAW.clear()
    cfg.PLATE_LENGTH_DM_RAW.update(saved_dm_raw)


def _run_optimization() -> tuple[list[dict], dict]:
    orders_2d = _build_merged_orders_2d()
    plate_order = AppPlateOrder.from_orders_2d(orders_2d)
    saved_details, saved_dm_raw = _push_plate_load_cfg(plate_order)
    service = OptimizationService()
    try:
        context = service.optimize(plate_order, orders_2d=orders_2d)
    finally:
        _restore_plate_load_cfg(saved_details, saved_dm_raw)
    result = context.optimization_result or {}
    assert result, "optimizer returned empty result"
    return orders_2d, result


def test_secondary_parent_refs_are_primaries_that_have_residual():
    """
    Каждый назначенный parent_instance_id должен указывать на первичный рез
    с rest > 0. Несколько вторичных строк могут делить одного родителя (pieces > 1).
    """
    _, result = _run_optimization()
    primary = result.get("primary_cuts") or []
    secondary = result.get("secondary_cuts") or []
    prim_with_rest = {
        p["primary_instance_id"]
        for p in primary
        if p.get("primary_instance_id") and int(p.get("rest") or 0) > 0
    }
    assigned_parents = [cut.get("parent_instance_id") for cut in secondary if cut.get("parent_instance_id")]
    assert assigned_parents, "expected at least one secondary with assigned parent"
    for pid in assigned_parents:
        assert pid in prim_with_rest, f"parent {pid} not in primaries with residual"


def test_secondary_orphan_keeps_source_geometry_and_order_coverage():
    """
    Если есть null-parent secondary — геометрия source должна находить кандидата
    в первичных остатках; покрытие заказа всегда полное.
    """
    orders_2d, result = _run_optimization()
    primary = result.get("primary_cuts") or []
    secondary = result.get("secondary_cuts") or []
    orphans = [cut for cut in secondary if not cut.get("parent_instance_id")]

    def _len_close(a: float, b: float) -> bool:
        return abs(float(a) - float(b)) < 0.06

    for orphan in orphans:
        source_rest = int(orphan.get("source") or 0)
        source_lengths = orphan.get("source_lengths") or []
        assert source_lengths, f"orphan secondary has no source_lengths: {orphan}"
        source_length = float(source_lengths[0])
        source_opt_ids = list(orphan.get("source_opt_ids") or [])
        assert source_opt_ids, f"orphan secondary has no source_opt_ids: {orphan}"

        residual_candidates = [
            prim
            for prim in primary
            if int(prim.get("rest") or 0) == source_rest
            and int(prim.get("rest") or 0) > 0
            and _len_close(float((prim.get("lengths") or [0])[0]), source_length)
        ]
        assert residual_candidates, "expected matching primary residual geometry for orphan"

    demand_2d: dict[tuple[float, int, int], int] = {}
    for row in orders_2d:
        key = (round(float(row["length"]), 2), int(row["width"]), _norm_lc(row.get("load_code", 8)))
        demand_2d[key] = demand_2d.get(key, 0) + int(row["qty"])
    coverage = verify_coverage(demand_2d, primary, secondary)
    assert coverage["ok"], coverage


def test_secondary_parent_assignment_emits_warning_with_missing_count(caplog: pytest.LogCaptureFixture):
    """
    При неполном назначении parent (null-parent) — warning с recovered_by_geometry/missing.
    Если все вторички привязаны — блок диагностики не логируется.
    """
    caplog.set_level(logging.WARNING, logger="core.optimization")
    _orders_2d, result = _run_optimization()
    null_parent_count = sum(1 for cut in (result.get("secondary_cuts") or []) if not cut.get("parent_instance_id"))

    records = [rec for rec in caplog.records if "secondary parent assignment" in rec.getMessage()]
    if null_parent_count > 0:
        assert records, "expected warning log for secondary parent assignment diagnostics"
        last = records[-1]
        assert isinstance(last.args, tuple) and len(last.args) == 2
        recovered_by_geometry, missing = last.args
        assert int(missing) >= 1
        assert int(missing) <= null_parent_count
        assert int(recovered_by_geometry) >= 0
    else:
        assert not records


@pytest.mark.parametrize(
    ("qty", "pieces", "expected"),
    [
        (5, 2, [2, 2, 1]),
        (4, 2, [2, 2]),
        (3, 1, [1, 1, 1]),
        (7, 3, [3, 3, 1]),
    ],
)
def test_batch_sizes_for_secondary_z_sec(qty: int, pieces: int, expected: list[int]) -> None:
    """Один родительский остаток на батч длиной <= pieces (согласовано с cap_sec ILP)."""
    assert _batch_sizes_for_secondary_z_sec(qty, pieces) == expected
    assert sum(expected) == qty
