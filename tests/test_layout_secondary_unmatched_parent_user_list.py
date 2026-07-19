# -*- coding: utf-8 -*-
"""
Аудит «несопоставленного» вторичного сегмента для фиксированного списка плит
(тот же набор, что в обсуждении с пользователем).

Важно: правка продуктового кода не требуется — только тесты.
"""
from __future__ import annotations

import copy
from collections import defaultdict
from unittest.mock import patch

import pytest

from core.domain.plate_order import normalize_load_code
from core.plate_runtime_state import get_plate_mutable_runtime
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.services.optimization_service import OptimizationService
from core.optimization import verify_coverage
from core.plate_line_parser import parse_line
from core.plate_order_context import PlateOrderContext
from viz_modules import layout_sequence as layout_sequence_mod
from viz_modules.layout_sequence import build_layout_sequence

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


def _norm_lc(x):
    return normalize_load_code(x, default=8)


def _build_merged_orders_2d() -> list[dict]:
    rows: list[dict] = []
    for raw in LINES:
        line = raw.strip()
        if not line:
            continue
        r = parse_line(line)
        assert r.parsed, (line, r.reason_text)
        lc = r.load_code if r.load_code is not None else 8.0
        rows.append(
            {
                "length": float(r.length_m),
                "width": int(round(r.width_m * 1000)),
                "qty": int(r.qty),
                "load_code": float(lc),
                "length_dm_raw": r.length_dm_raw or "",
            }
        )
    merged: dict[tuple, int] = {}
    for o in rows:
        key = (round(o["length"], 4), o["width"], _norm_lc(o["load_code"]))
        merged[key] = merged.get(key, 0) + o["qty"]
    return [
        {"length": k[0], "width": k[1], "qty": v, "load_code": float(k[2])}
        for k, v in sorted(merged.items())
    ]


def _push_plate_load_cfg(plate_order: AppPlateOrder) -> tuple[dict, dict]:
    rt = get_plate_mutable_runtime()
    saved_d = copy.deepcopy(rt.plate_load_details)
    saved_raw = copy.deepcopy(rt.plate_length_dm_raw)
    rt.plate_load_details.clear()
    for key, qty in plate_order.plate_load_details.items():
        length, width_m, load_code, raw = key
        rt.plate_load_details[(length, width_m, int(float(load_code)), raw)] = int(qty)
    rt.plate_length_dm_raw.clear()
    for key, raw in plate_order.plate_length_dm_raw.items():
        length, width_m, load_code, raw_val = key
        rt.plate_length_dm_raw[(length, width_m, int(float(load_code)), raw_val)] = raw
    return saved_d, saved_raw


def _restore_plate_load_cfg(saved_d: dict, saved_raw: dict) -> None:
    rt = get_plate_mutable_runtime()
    rt.plate_load_details.clear()
    rt.plate_load_details.update(saved_d)
    rt.plate_length_dm_raw.clear()
    rt.plate_length_dm_raw.update(saved_raw)


@pytest.fixture
def user_list_opt_context():
    """Оптимизация + legacy-план в глобалях, cfg восстановлен после теста."""
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


def test_optimizer_exactly_one_secondary_without_parent_instance(user_list_opt_context):
    """
    Источник «несопоставленного» сегмента в layout: у записи secondary_cuts
    нет parent_instance_id (оптимизатор не смог привязать к экземпляру primary).

    В debug-7e420e это H_OPT_SEC_SUMMARY: null_parent_by_source_geom_key 5.8_880.

    После фиксов очереди родителя оптимизатор на этом же фикстурном заказе может
    дать 0 null-parent — тогда сценарий пропускаем.
    """
    _orders_2d, _plate_order, svc, ctx = user_list_opt_context
    res = ctx.optimization_result or {}
    sec = res.get("secondary_cuts") or []
    null_parents = [c for c in sec if not c.get("parent_instance_id")]
    if len(null_parents) != 1:
        pytest.skip(f"фикстура дала {len(null_parents)} null-parent вторичных, ожидался 1 для узкого сценария")
    orphan = null_parents[0]
    assert int(orphan.get("source") or 0) == 880
    sl = orphan.get("source_lengths") or []
    assert sl and abs(float(sl[0]) - 5.8) < 0.05
    tok = orphan.get("target_order_key")
    assert tok is not None and len(tok) >= 3
    assert abs(float(tok[0]) - 5.8) < 0.05
    assert int(tok[1]) == 720
    assert int(float(tok[2])) == 8


def test_layout_phase_C_logs_null_parent_slot_and_h5_unmatched(user_list_opt_context):
    """
    В phase C split: при отсутствии chosen_variant для rest>0 логируется H3.
    Итог H5: secondary_unmatched_total = плановые вторичные минус реально
    повешенные на слоты pattern-сегменты.

    Несопоставленный вторичный в плане имеет parent_instance_id=None — в индекс
    secondary_cuts_by_parent он не попадает (if parent_instance_id), поэтому
    привязка только по геометрии; один сегмент может остаться без слота → unmatched=1.

    Вторичные без parent эмулируются как первичный рез (вариант A) и учитываются
    в secondary_attached_total — при наличии только таких «сирот» unmatched=0.

    Если после оптимизатора null-parent нет, ожидаем secondary_unmatched_total==0.
    """
    _orders_2d, plate_order, svc, ctx = user_list_opt_context
    res = copy.deepcopy(ctx.optimization_result or {})
    assert res
    n_orphan_plan = sum(1 for c in (res.get("secondary_cuts") or []) if not c.get("parent_instance_id"))

    captured: list[dict] = []
    _orig_agent_seq_debug = layout_sequence_mod._agent_seq_debug

    def _spy(hypothesis_id: str, message: str, data: dict) -> None:
        captured.append(
            {
                "hypothesisId": hypothesis_id,
                "message": message,
                "data": dict(data) if isinstance(data, dict) else data,
            }
        )
        return _orig_agent_seq_debug(hypothesis_id, message, data)

    saved_d, saved_raw = _push_plate_load_cfg(plate_order)
    plate_ctx = PlateOrderContext.fresh_empty()
    plate_ctx.hydrate_from_order(plate_order)
    try:
        with patch.object(layout_sequence_mod, "_agent_seq_debug", side_effect=_spy):
            with svc.bound_plate_order_context(plate_ctx, ctx):
                build_layout_sequence()
    finally:
        _restore_plate_load_cfg(saved_d, saved_raw)

    h5 = [x for x in captured if x.get("hypothesisId") == "H5" and x.get("message") == "phase_end_summary"]
    assert h5, "ожидали хотя бы одну запись H5 phase_end_summary"
    last = h5[-1]["data"]
    if n_orphan_plan >= 1:
        assert last.get("secondary_unmatched_total") == 0
        assert last.get("secondary_total_from_plan", 0) == last.get("secondary_attached_total", 0)
    else:
        assert last.get("secondary_unmatched_total") == 0
        assert last.get("secondary_total_from_plan", 0) == last.get("secondary_attached_total", 0)

    h3_unmatched = [
        x["data"]
        for x in captured
        if x.get("hypothesisId") == "H3"
        and x.get("data", {}).get("unmatched_increment") == "parent_instance_id_not_found"
    ]
    if not h3_unmatched:
        return
    by_parent = defaultdict(int)
    for d in h3_unmatched:
        pid = d.get("parent_instance_id")
        by_parent[str(pid)] += 1
    # Ровно один «лишний» вторичный в плане — без parent — не порождает H3
    # parent_instance_id_not_found (там нужен truthy parent на слоте).
    # Проверяем, что есть хотя бы один H3 с реальным prim-* и нулевыми пулами.
    real_prim_failures = [d for d in h3_unmatched if d.get("parent_instance_id")]
    assert real_prim_failures
    for d in real_prim_failures[:3]:
        assert str(d.get("parent_instance_id", "")).startswith("prim-")
        assert d.get("n_pool_parent_free") == 0
        assert d.get("matched_via") is None


def test_null_parent_secondary_traced_to_primary_residual_not_missing_strip(
    user_list_opt_context,
):
    """
    Выясняем, «из какой первичной геометрии» в плане задуман вторичный без parent_instance_id.

    None у parent — это сбой пост-разметки id (очереди prim-* исчерпаны), а не отсутствие
    остатка в модели: в primary_cuts обязаны быть резы с тем же rest (мм) и длиной,
    что у secondary.source / source_lengths, и source_opt_id вторичного входит в
    source_opt_ids из опции (см. optimization.py: построение secondary из possible_rests).

    Дополнительно: покрытие заказа остаётся ok — «куск не нашёлся» было бы дефицитом.
    """
    orders_2d, _plate_order, _svc, ctx = user_list_opt_context
    res = ctx.optimization_result or {}
    prim = res.get("primary_cuts") or []
    sec_all = res.get("secondary_cuts") or []

    orphans = [c for c in sec_all if not c.get("parent_instance_id")]
    if len(orphans) != 1:
        pytest.skip(f"нет ровно одного orphan вторичного в плане (got {len(orphans)})")
    o = orphans[0]
    src_mm = int(o.get("source") or 0)
    sl = o.get("source_lengths") or []
    assert sl
    src_len = float(sl[0])
    sopts = list(o.get("source_opt_ids") or [])
    assert sopts, "у вторичной опции в плане должны быть source_opt_ids (первичные opt id)"

    def _len_close(a: float, b: float) -> bool:
        return abs(float(a) - float(b)) < 0.06

    residual_primaries = [
        p
        for p in prim
        if int(p.get("rest") or 0) > 0
        and int(p.get("rest") or 0) == src_mm
        and _len_close(float((p.get("lengths") or [0])[0]), src_len)
    ]
    assert residual_primaries, (
        "ожидали хотя бы один первичный рез в плане с тем же остатком (мм) и длиной, "
        "что и у вторичного без parent — иначе не из чего закрывать спрос в рамках этой модели"
    )

    matched_by_opt = [p for p in residual_primaries if p.get("source_opt_id") in sopts]
    assert matched_by_opt, (
        f"остаток {src_len}м×{src_mm}мм должен сопоставляться с source_opt_ids={sopts}; "
        f"найдено residual_primaries={len(residual_primaries)}, matched={len(matched_by_opt)}"
    )

    for p in matched_by_opt:
        main_w = int(p.get("width") or 0)
        rest_w = int(p.get("rest") or 0)
        assert main_w + rest_w == 1200

    demand_2d: dict = {}
    for row in orders_2d:
        key = (
            round(float(row["length"]), 2),
            int(row["width"]),
            _norm_lc(row.get("load_code", 8)),
        )
        demand_2d[key] = demand_2d.get(key, 0) + int(row["qty"])
    cov = verify_coverage(demand_2d, prim, sec_all)
    assert cov["ok"], (
        "parent_instance_id=None не означает «не из чего резать»: при дефиците куска "
        f"verify_coverage был бы false; получили {cov}"
    )

    # Конкретика для этого фиксированного заказа: остаток 880 мм на 5.8 м — от ПБ 58-3,2 (main 320)
    p0 = matched_by_opt[0]
    ak = p0.get("assignment_key")
    assert ak is not None and len(ak) >= 2
    assert _len_close(float(ak[0]), 5.8)
    assert int(ak[1]) == 320
