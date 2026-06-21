# -*- coding: utf-8 -*-
"""
Крупный заказ: вторичные без parent_instance_id — трассировка к первичной геометрии
и проверка, что None ≠ «нет куска» (verify_coverage).

Код приложения не меняется.
"""
from __future__ import annotations

import copy

import pytest

from core.domain.plate_order import normalize_load_code
from core.plate_runtime_state import get_plate_mutable_runtime
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.services.optimization_service import OptimizationService
from core.optimization import verify_coverage
from core.plate_line_parser import parse_line

# Список из запроса пользователя (как в исходном тексте заказа).
LINES_RAW = r"""
Плиты ПБ 69-12-8п    20
Плиты ПБ 45-10,8-8п    4
Плиты ПБ 45-12-6п    12
Плиты ПБ 45-7,0-6п    4
Плиты ПБ 42-12-6п    8
Плиты ПБ 42-5,3-6п    4
Плиты ПБ 42,6-12-8п    5
Плиты ПБ 42,6-12-10п    3
Плиты ПБ 63,9-12-8п    6
Плиты ПБ 63,9-12-10п    2
Плиты ПБ 69,3-12-8п    3
Плиты ПБ 69,3-12-10п    3
Плиты ПБ 37,2-12-8п    3
Плиты ПБ 37,2-12-10п    3
Плиты ПБ 42,6-5,3-10п    1
Плиты ПБ 63,9-5,3-10п    1
Плиты ПБ 37-12-8п    8
Плиты ПБ 73-12-8п    4
Плиты ПБ 38-12-8п    4
Плиты ПБ 56-12-8п    2
Плиты ПБ 58,5-12-8п    1
Плиты ПБ 54,5-12-8п    5
Плиты ПБ 59,8-12-8п    6
Плиты ПБ 63,5-12-8п    2
Плиты ПБ 42,8-12-8п    1
Плиты ПБ 37-12-8п    2
Плиты ПБ 77-12-8п    5
Плиты ПБ 52-12-8п    3
Плиты ПБ 38-12-8п    2
Плиты ПБ 42-12-8п    1
Плиты ПБ 43-12-8п    5
Плиты ПБ 57-12-8п    1
Плиты ПБ 58-12-8п    3
Плиты ПБ 69-12-8п    6
Плиты ПБ 69-12-8п    14
Плиты ПБ 25-12-8п    3
Плиты ПБ 80-12-8п    7
Плиты ПБ 80-5,3-8п    1
Плиты ПБ 80-6,65-8п    1
Плиты ПБ 60-5,3-8п    2
Плиты ПБ 60-6,65-8п    2
Плиты ПБ 54,3-5,3-8п    6
Плиты ПБ 51-3,2-8п    2
Плиты ПБ 49,9-6,65-8п    1
Плиты ПБ 40-12-8п    4
Плиты ПБ 40-9,2-8п    1
Плиты ПБ 42-12-8п    2
Плиты ПБ 42-9,0-8п    1
Плиты ПБ 42-3,0-8п    1
Плиты ПБ 33-12-8п    1
Плиты ПБ 38,6-12-8п    7
Плиты ПБ 37,9-12-8п    1
Плиты ПБ 37,9-10,2-8п    1
Плиты ПБ 37,9-9,2-8п    2
Плиты ПБ 37,9-9,0-8п    1
Плиты ПБ 37,9-3,2-8п    1
Плиты ПБ 39,2-12-8п    1
Плиты ПБ 89-12-8п    9
Плиты ПБ 42-12-8п    1
Плиты ПБ 77-12-8п    2
Плиты ПБ 77-10,2-8п кос    1
Плиты ПБ 71-12-8п кос. рез    1
Плиты ПБ 57-12-8п    1
Плиты ПБ 70,5-12-8п    16
Плиты ПБ 33,9-10,8-8п    2
Плиты ПБ 49,7-12-8п    12
Плиты ПБ 49,7-10,8-8п    2
Плиты ПБ 34-12-8п    2
Плиты ПБ 60-10,2-8п    3
Плиты ПБ 61,8-12-8п    9
Плиты ПБ 60-12-8п    15
Плиты ПБ 57,1-12-8п    19
Плиты ПБ 31,1-12-8п    4
Плиты ПБ 61,8-5,0-8п    1
Плиты ПБ 60-5,3-8п    1
Плиты ПБ 61,8-12-8 1 выб    1
Плиты ПБ 61,8-12-8 1 выб    1
Плиты ПБ 61,8-12-8 2 выб    1
Плиты ПБ 60-12-8 2 выб    1
Плиты ПБ 60-12-8 2 выб    1
Плиты ПБ 57,1-12-8 1 выб    1
Плиты ПБ 57,1-12-8 1 выб    1
Плиты ПБ 57,1-12-8 1 выб    1
Плиты ПБ 57,1-12-8 1 выб    1
Плиты ПБ 31,1-12-8 1 выб    1
Плиты ПБ 58-12-12,5п    26
Плиты ПБ 62-12-12,5п    2
Плиты ПБ 58-12-12,5 2 выб    12
Плиты ПБ 80-12-8п    7
Плиты ПБ 80-5,3-8п    1
Плиты ПБ 80-6,65-8п    1
Плиты ПБ 60-5,3-8п    1
Плиты ПБ 60-6,65-8п    2
Плиты ПБ 54,3-5,3-8п    6
Плиты ПБ 51-3,2-8п    2
Плиты ПБ 49,9-6,65-8п    1
Плиты ПБ 49,9-6,65-8п    1
Плиты ПБ 40,3-12-8п    5
Плиты ПБ 40,3-2,6-8п    1
Плиты ПБ 42,2-12-8п    3
Плиты ПБ 42,2-2,6-8п    1
Плиты ПБ 25,2-12-8п    2
Плиты ПБ 26,7-6,65-8п    1
Плиты ПБ 45-12-8п    10
Плиты ПБ 59,5-12-8п    2
Плиты ПБ 39,0-12-8п    6
Плиты ПБ 20,6-12-8п    13
Плиты ПБ 20,6-7,2-8п    3
Плиты ПБ 25,4-12-8п    8
Плиты ПБ 25,4-3,0-8п    4
Плиты ПБ 24,4-7,2-8п    4
Плиты ПБ 40,0-12-8п    6
Плиты ПБ 74,1-12-8п    40
Плиты ПБ 63,7-5,3-8п    8
Плиты ПБ 74,1-9,2-8п    4
Плиты ПБ 63,7-12-8п    51
Плиты ПБ 63,7-8,6-8п    2
""".strip()


def _norm_lc(x):
    return normalize_load_code(x, default=8)


def _build_merged_orders_2d() -> list[dict]:
    rows: list[dict] = []
    for raw in LINES_RAW.splitlines():
        line = raw.strip()
        if not line:
            continue
        r = parse_line(line)
        assert r.parsed, f"строка не разобрана: {line!r} → {r.reason_text}"
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


@pytest.fixture(scope="module")
def large_order_opt_result():
    orders_2d = _build_merged_orders_2d()
    plate_order = AppPlateOrder.from_orders_2d(orders_2d)
    saved_d, saved_raw = _push_plate_load_cfg(plate_order)
    svc = OptimizationService()
    try:
        ctx = svc.optimize(plate_order, orders_2d=orders_2d)
    finally:
        _restore_plate_load_cfg(saved_d, saved_raw)
    res = ctx.optimization_result
    assert res, "оптимизация должна вернуть результат"
    return orders_2d, res


def test_large_order_every_line_parses():
    """Все строки списка успешно разбираются parse_line."""
    n = 0
    for raw in LINES_RAW.splitlines():
        line = raw.strip()
        if not line:
            continue
        r = parse_line(line)
        assert r.parsed, (line, r.reason_text)
        n += 1
    assert n >= 50


def test_large_order_verify_coverage_ok(large_order_opt_result):
    """Покрытие спроса: вторичные без parent не означают дыру в заказе."""
    orders_2d, res = large_order_opt_result
    prim = res.get("primary_cuts") or []
    sec = res.get("secondary_cuts") or []
    demand_2d: dict = {}
    for row in orders_2d:
        key = (
            round(float(row["length"]), 2),
            int(row["width"]),
            _norm_lc(row.get("load_code", 8)),
        )
        demand_2d[key] = demand_2d.get(key, 0) + int(row["qty"])
    cov = verify_coverage(demand_2d, prim, sec)
    assert cov["ok"], cov


def test_large_order_null_parent_secondaries_trace_to_primary_residual(
    large_order_opt_result,
):
    """
    Для каждого secondary без parent_instance_id:
    - в primary_cuts есть резы с тем же (длина, rest мм), что source_lengths/source
      (физический «кусок» в плане есть — None у parent ≠ «не из чего резать»);
    - если у первичного задан source_opt_id — он может входить в source_opt_ids
      вторичной опции; иначе это пост-коррекция / force-add (source_opt_id None),
      но геометрия остатка всё равно согласована с полосой 1200 мм.

    Если вторичных без parent нет — тест пропускается.
    """
    orders_2d, res = large_order_opt_result
    prim = res.get("primary_cuts") or []
    sec_all = res.get("secondary_cuts") or []

    orphans = [c for c in sec_all if not c.get("parent_instance_id")]
    if not orphans:
        pytest.skip("нет secondary без parent_instance_id — трассировка не требуется")

    def _len_close(a: float, b: float) -> bool:
        return abs(float(a) - float(b)) < 0.08

    for o in orphans:
        src_mm = int(o.get("source") or 0)
        sl = o.get("source_lengths") or []
        assert sl, f"orphan без source_lengths: {o}"
        src_len = float(sl[0])
        sopts = list(o.get("source_opt_ids") or [])
        assert sopts, f"orphan без source_opt_ids: {o}"

        residual_primaries = [
            p
            for p in prim
            if int(p.get("rest") or 0) > 0
            and int(p.get("rest") or 0) == src_mm
            and _len_close(float((p.get("lengths") or [0])[0]), src_len)
        ]
        assert residual_primaries, (
            f"нет первичного остатка {src_len}м×{src_mm}мм для orphan target={o.get('target_order_key')}"
        )

        matched = [p for p in residual_primaries if p.get("source_opt_id") in sopts]
        chosen = matched if matched else residual_primaries

        for p in chosen:
            main_w = int(p.get("width") or 0)
            rest_w = int(p.get("rest") or 0)
            assert main_w + rest_w == 1200, (p, o)

        if matched:
            for p in matched:
                assert p.get("source_opt_id") in sopts

    demand_2d: dict = {}
    for row in orders_2d:
        key = (
            round(float(row["length"]), 2),
            int(row["width"]),
            _norm_lc(row.get("load_code", 8)),
        )
        demand_2d[key] = demand_2d.get(key, 0) + int(row["qty"])
    assert verify_coverage(demand_2d, prim, sec_all)["ok"]


def test_large_order_reports_null_parent_count(large_order_opt_result):
    """Документируем факт: сколько вторичных без parent (может быть 0)."""
    _orders_2d, res = large_order_opt_result
    sec = res.get("secondary_cuts") or []
    n_null = sum(1 for c in sec if not c.get("parent_instance_id"))
    # Не assert на конкретное число — только информативная нижняя граница
    assert n_null >= 0
    assert len(sec) >= n_null


def test_large_order_orphans_intended_primary_strips_via_assignment_key(
    large_order_opt_result,
):
    """
    «Из какой плиты задуман» вторичный без parent_instance_id:

    В плане первичного реза каждая строка с rest>0 — физический остаток полосы.
    Поле assignment_key — это спрос (L, ширина мм, класс), который закрывает
    **основная** полоса той же операции (width + rest = 1200). Геометрия
    остатка совпадает с secondary.source / source_lengths.

    parent_instance_id=None возникает, когда очередь экземпляров prim-* по
    source_ids опустела, а fallback по геометрии не успел/не смог выдать id
    (см. core/optimization.py): это не то же самое, что «в модели нет куска» —
    на то указывают найденные здесь primary и verify_coverage в соседнем тесте.
    """
    orders_2d, res = large_order_opt_result
    prim = res.get("primary_cuts") or []
    sec_all = res.get("secondary_cuts") or []
    orphans = [c for c in sec_all if not c.get("parent_instance_id")]
    if not orphans:
        pytest.skip("нет secondary без parent_instance_id")

    def _len_close(a: float, b: float) -> bool:
        return abs(float(a) - float(b)) < 0.08

    demand_keys = {
        (round(float(row["length"]), 2), int(row["width"]), _norm_lc(row.get("load_code", 8)))
        for row in orders_2d
    }

    for o in orphans:
        src_mm = int(o.get("source") or 0)
        sl = o.get("source_lengths") or []
        assert sl
        src_len = float(sl[0])
        sopts = list(o.get("source_opt_ids") or [])
        assert sopts

        residual = [
            p
            for p in prim
            if int(p.get("rest") or 0) == src_mm
            and int(p.get("rest") or 0) > 0
            and _len_close(float((p.get("lengths") or [0])[0]), src_len)
        ]
        assert residual

        matched = [p for p in residual if p.get("source_opt_id") in sopts]
        chosen = matched if matched else residual

        aks = []
        for p in chosen:
            ak = p.get("assignment_key")
            assert ak is not None and len(ak) >= 3
            L_d, W_d, lc_d = float(ak[0]), int(ak[1]), _norm_lc(ak[2])
            assert (round(L_d, 2), W_d, lc_d) in demand_keys
            main_w = int(p.get("width") or 0)
            rest_w = int(p.get("rest") or 0)
            assert main_w + rest_w == 1200
            assert W_d == int(p.get("demand_width") or W_d)
            aks.append((ak, main_w, rest_w))

        assert aks, "ожидали хотя бы один assignment_key для трассировки orphan"
