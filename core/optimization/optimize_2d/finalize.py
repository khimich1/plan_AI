# -*- coding: utf-8 -*-
"""Phase C: PlateAudit checkpoints, post-correction, coverage verify, slot attribution."""

from __future__ import annotations

from collections import Counter

from core.domain.plate_order import normalize_load_code
from core.optimization.coverage_verify import verify_coverage
from core.optimization.debug_log import (
    _DEBUG_LOG_5b5324,
    _DEBUG_LOG_COMMON,
    _dbg_open_append,
    _opt_debug_enabled,
)
from core.optimization.optimization_debug_impl import (
    _DEBUG_AGENT_LOG_EBB546,
    _DEBUG_LOG_2D5C43,
    _DEBUG_LOG_EF42AE,
    _debug_runtime_write_648532,
)
from core.optimization.optimize_2d.state import norm_demand_key
from core.optimization.order_dispatch import _next_slot_info
from core.optimization.result_contract import opt_ok


def run_two_d_phase_finalize(
    *,
    demand_2d: dict,
    plate_width: int,
    slot_lists: dict,
    slot_cursors: dict,
    no_sources_keys,
    solver_status: str,
    audit,
    result: dict,
    n_solid_primary_plates: int,
    n_cut_primary_plates: int,
    next_primary_instance_id: int,
) -> dict:
    """
    Post-solver: audit checkpoints, demand post-correction, verify_coverage,
    primary/secondary slot attribution via _next_slot_info, plate_assignments.
    Mutates result, slot_cursors, and audit in place.
    """
    _next_primary_instance_id = next_primary_instance_id

    # #region agent log
    if _opt_debug_enabled():
        try:
            import json as _agent_json
            import time as _agent_time

            with _dbg_open_append(_DEBUG_AGENT_LOG_EBB546) as _agent_f:
                _agent_f.write(
                    _agent_json.dumps(
                        {
                            "sessionId": "ebb546",
                            "runId": "solver-localization",
                            "hypothesisId": "O6",
                            "location": "core/optimization/optimize_2d/finalize.py:solver_output_pre",
                            "message": "Фактические primary/secondary перед PlateAudit solver_output",
                            "data": {
                                "primary_cuts_len": len(result.get("primary_cuts") or []),
                                "secondary_cuts_len": len(result.get("secondary_cuts") or []),
                                "total_for_audit": len(result.get("primary_cuts") or [])
                                + len(result.get("secondary_cuts") or []),
                                "secondary_sample": [
                                    {
                                        "source": c.get("source"),
                                        "cuts": c.get("cuts"),
                                        "lengths": c.get("lengths"),
                                        "target_order_key": c.get("target_order_key"),
                                        "load_code": c.get("load_code"),
                                    }
                                    for c in (result.get("secondary_cuts") or [])[:20]
                                ],
                            },
                            "timestamp": int(_agent_time.time() * 1000),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion

    # PlateAudit: checkpoint после сбора данных из solver (до пост-коррекции)
    audit.checkpoint("solver_output", result["primary_cuts"] + result["secondary_cuts"])

    # Универсальная пост-коррекция: считаем уже покрытый спрос по ОБОИМ источникам,
    # а добираем только реально unmet demand.
    # BUG-3 FIX: нормализуем ключи через norm_demand_key, чтобы load_code 8/800 не давал
    # ложный «have=0» и не порождал дубликаты или пропуски.
    planned_coverage_by_key: Counter[tuple] = Counter()
    for cut in result["primary_cuts"]:
        assignment_key = cut.get("assignment_key")
        if assignment_key:
            planned_coverage_by_key[norm_demand_key(assignment_key)] += 1
    for cut in result["secondary_cuts"]:
        target_key = cut.get("target_order_key")
        if target_key:
            planned_coverage_by_key[norm_demand_key(target_key)] += 1

    _post_correction_added = 0
    for (L, W, lc), need in demand_2d.items():
        have = planned_coverage_by_key.get(norm_demand_key((L, W, lc)), 0)
        if have >= need:
            continue
        if W > plate_width:
            continue
        rest = (plate_width - W) if W < plate_width else 0
        for _ in range(need - have):
            primary_instance_id = f"prim-{_next_primary_instance_id}"
            _next_primary_instance_id += 1
            result["primary_cuts"].append(
                {
                    "width": W,
                    "demand_width": W,
                    "rest": rest,
                    "qty": 1,
                    "lengths": [L],
                    "load_code": lc,
                    "assignment_key": (L, W, lc),
                    "identity_match_type": "post_correction_pending",
                    "source_opt_id": None,
                    "primary_instance_id": primary_instance_id,
                }
            )
            result["total_plates"] += 1
            _post_correction_added += 1
    if _post_correction_added:
        import logging as _logging

        _logging.getLogger(__name__).warning(
            "[OPT_2D] [POST_CORRECTION] Добавлено %d плит восстановлением покрытия спроса.",
            _post_correction_added,
        )

    # BUG-1 FIX: явная страховочная проверка для no_sources_keys.
    # Плиты без источников в модели должны быть добавлены как «прямой рез»,
    # если пост-коррекция их не покрыла (например, из-за W > plate_width скипа).
    if no_sources_keys:
        import logging as _logging

        _logger_ns = _logging.getLogger(__name__)
        # Пересчитываем покрытие с учётом только что добавленных пост-коррекцией
        _final_coverage: Counter[tuple] = Counter()
        for _cut in result["primary_cuts"]:
            _ak = _cut.get("assignment_key")
            if _ak:
                _final_coverage[norm_demand_key(_ak)] += 1
        for _cut in result["secondary_cuts"]:
            _tk = _cut.get("target_order_key")
            if _tk:
                _final_coverage[norm_demand_key(_tk)] += 1

        _force_added = 0
        for (_key, _qty) in no_sources_keys:
            _L, _W, _lc = _key
            _have = _final_coverage.get(norm_demand_key(_key), 0)
            if _have >= _qty:
                continue
            _shortfall = _qty - _have
            _rest = max(0, plate_width - _W) if _W <= plate_width else 0
            for _ in range(_shortfall):
                primary_instance_id = f"prim-{_next_primary_instance_id}"
                _next_primary_instance_id += 1
                result["primary_cuts"].append(
                    {
                        "width": _W,
                        "demand_width": _W,
                        "rest": _rest,
                        "qty": 1,
                        "lengths": [_L],
                        "load_code": _lc,
                        "assignment_key": (_L, _W, _lc),
                        "identity_match_type": "force_added_no_sources",
                        "source_opt_id": None,
                        "primary_instance_id": primary_instance_id,
                    }
                )
                result["total_plates"] += 1
                _force_added += 1
                _final_coverage[norm_demand_key(_key)] += 1
        if _force_added:
            _logger_ns.warning(
                "[OPT_2D] [BUG-1 FIX] Force-added %d плит из no_sources_keys "
                "(не было источника в оптимизаторе, добавлены как прямой рез).",
                _force_added,
            )

    # PlateAudit: checkpoint после пост-коррекции и force-add (финальный срез оптимизатора)
    audit.checkpoint("post_correction", result["primary_cuts"] + result["secondary_cuts"])
    if audit.has_losses("demand_2d", "post_correction"):
        import logging as _log_mod2

        _log_mod2.getLogger(__name__).error(
            "[AUDIT] Потери плит в оптимизаторе!\n%s", audit.summary()
        )
    else:
        import logging as _log_mod2

        _log_mod2.getLogger(__name__).info("[AUDIT] Оптимизатор: потерь нет.\n%s", audit.summary())
    # #region agent log
    if _opt_debug_enabled():
        try:
            _missing_after_post = {}
            for (_L, _W, _lc), _need in demand_2d.items():
                _have = planned_coverage_by_key.get(norm_demand_key((_L, _W, _lc)), 0)
                if _have < _need:
                    _missing_after_post[str([_L, _W, _lc])] = int(_need - _have)
            _debug_runtime_write_648532(
                "run1",
                "H3_post_correction_balance",
                "core/optimization/optimize_2d/finalize.py:post_correction",
                "Demand minus planned coverage after post-correction",
                {
                    "demand_total": int(sum(demand_2d.values())),
                    "planned_coverage_total": int(sum(planned_coverage_by_key.values())),
                    "missing_keys_after_post": _missing_after_post,
                    "missing_total_after_post": int(sum(_missing_after_post.values())),
                },
            )
        except Exception:
            pass
    # #endregion
    result["_plate_audit"] = audit

    # ОБЯЗАТЕЛЬНАЯ ПРОВЕРКА ПОКРЫТИЯ: логируем итог по demand_2d vs primary+secondary.
    # Сейчас работает как наблюдатель; в этап 2 модель должна выдавать ok=True всегда.
    _coverage_summary = verify_coverage(
        demand_2d, result["primary_cuts"], result["secondary_cuts"]
    )
    result["_coverage_summary"] = _coverage_summary
    import logging as _log_cov

    _cov_logger = _log_cov.getLogger(__name__)
    if not _coverage_summary["ok"]:
        _cov_logger.error(
            "[OPT_2D] [COVERAGE] missing=%d плит по %d ключам, surplus=%d плит по %d ключам",
            sum(_coverage_summary["missing"].values()),
            len(_coverage_summary["missing"]),
            sum(_coverage_summary["surplus"].values()),
            len(_coverage_summary["surplus"]),
        )
    else:
        _cov_logger.info(
            "[OPT_2D] [COVERAGE] OK: demand=%d, covered=%d, surplus=%d по %d ключам",
            _coverage_summary["demand_total"],
            _coverage_summary["covered_total"],
            sum(_coverage_summary["surplus"].values()),
            len(_coverage_summary["surplus"]),
        )

    # #region agent log (2d5c43) H1,H2,H5: demand vs primary_cuts, 6m 530/1200
    if _opt_debug_enabled():
        try:
            _log_2d5c43 = _DEBUG_LOG_2D5C43
            _demand_total = sum(demand_2d.values())
            _target_keys = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
            _demand_by_key = {}
            for k, q in demand_2d.items():
                for tk in _target_keys:
                    if (
                        abs(round(k[0], 2) - round(tk[0], 2)) <= 0.02
                        and k[1] == tk[1]
                        and (k[2] == tk[2] or k[2] == "8")
                    ):
                        _demand_by_key[tuple(tk)] = _demand_by_key.get(tuple(tk), 0) + q
                        break
            _prim_total = len(result["primary_cuts"])
            _prim_by_key = {tk: 0 for tk in _target_keys}
            _prim_6_530_1200 = []
            for c in result["primary_cuts"]:
                L = round((c.get("lengths") or [0])[0], 2)
                W = c.get("demand_width", c.get("width", 0))
                lc = c.get("load_code", 8)
                for tk in _target_keys:
                    if (
                        abs(L - tk[0]) <= 0.02
                        and W == tk[1]
                        and (lc == tk[2] or lc == "8")
                    ):
                        _prim_by_key[tk] = _prim_by_key.get(tk, 0) + 1
                        break
                if 5.98 <= L <= 6.02 and W in (530, 1200) and len(_prim_6_530_1200) < 25:
                    _prim_6_530_1200.append(
                        {
                            "length": L,
                            "width": W,
                            "rest": c.get("rest", 0),
                            "plate_name": (c.get("plate_name") or "")[:60],
                        }
                    )
            _demand_by_key_ser = [list(k) + [v] for k, v in _demand_by_key.items()]
            _prim_by_key_ser = [list(k) + [v] for k, v in _prim_by_key.items()]
            with _dbg_open_append(_log_2d5c43) as _f:
                _f.write(
                    __import__("json").dumps(
                        {
                            "sessionId": "2d5c43",
                            "hypothesisId": "H1_H2_H5",
                            "location": "core/optimization/optimize_2d/finalize.py:h1_h5",
                            "message": "demand vs primary_cuts",
                            "data": {
                                "demand_total": _demand_total,
                                "demand_by_key": _demand_by_key_ser,
                                "primary_total": _prim_total,
                                "primary_by_key": _prim_by_key_ser,
                                "primary_6m_530_1200_sample": _prim_6_530_1200,
                            },
                            "timestamp": __import__("time").time(),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion
    print(f"[OPT_2D] ✓ Целых плит в начале: {n_solid_primary_plates}")
    print(f"[OPT_2D] ✓ Плит с резом (сгруппировано): {n_cut_primary_plates}")

    # Финальная атрибуция primary: теперь расходуем общий slot-ledger строго по demand key.
    _empty_primary_keys = []

    # Пересоздаём plate_assignments в правильном порядке
    # ИСПРАВЛЕНИЕ: добавляем load_code, иначе в треках подставляется 8п и списание ищет 8п вместо 10п
    result["plate_assignments"] = []
    for cut in result["primary_cuts"]:
        assignment_key = cut.get("assignment_key") or (
            (cut.get("lengths") or [0])[0],
            cut.get("demand_width", cut.get("width")),
            cut.get("load_code", 800),
        )
        plate_info = _next_slot_info(slot_lists, slot_cursors, assignment_key)
        if not plate_info:
            _empty_primary_keys.append(assignment_key)

        cut["load_code"] = (
            plate_info.get("load_code", cut.get("load_code", 800)) if plate_info else cut.get("load_code", 800)
        )
        cut["kp_id"] = plate_info.get("kp_id") if plate_info else None
        cut["customer"] = plate_info.get("customer") if plate_info else None
        cut["kp_date"] = plate_info.get("kp_date") if plate_info else None
        cut["plate_name"] = plate_info.get("plate_name") if plate_info else None
        cut["identity_match_type"] = plate_info.get("identity_match_type") if plate_info else "slot_exhausted"
        cut["concrete_grade"] = plate_info.get("concrete_grade") if plate_info else None

        for length in cut["lengths"]:
            result["plate_assignments"].append(
                {
                    "length": length,
                    "width": cut["width"],
                    "source": "primary",
                    "rest_width": cut["rest"],
                    "kp_id": cut.get("kp_id"),
                    "customer": cut.get("customer"),
                    "kp_date": cut.get("kp_date"),
                    "plate_name": cut.get("plate_name"),
                    "load_code": cut.get("load_code", 800),
                    "identity_match_type": cut.get("identity_match_type"),
                    "concrete_grade": cut.get("concrete_grade"),
                    "unit_id": cut.get("primary_instance_id"),
                    "parent_unit_id": None,
                    "source_opt_id": cut.get("source_opt_id"),
                }
            )

            # Сохраняем информацию об остатке для отслеживания
            if cut["rest"] > 0:
                result["rests_created"].append(
                    {
                        "length": length,
                        "rest_width_mm": cut["rest"],
                        "source_width_mm": cut["width"],
                    }
                )
    # #region agent log: summary empty plate_info (H2)
    if _empty_primary_keys and _opt_debug_enabled():
        try:
            _c = Counter(_empty_primary_keys)
            _summary = [{"key": list(k), "count": _c[k]} for k in sorted(_c.keys())]
            with _dbg_open_append(_DEBUG_LOG_COMMON) as _f:
                _f.write(
                    __import__("json").dumps(
                        {
                            "hypothesisId": "H2",
                            "location": "core/optimization/optimize_2d/finalize.py:empty_primary",
                            "message": "plan keys with empty plate_info (summary)",
                            "data": {
                                "by_key": _summary,
                                "total_plates_empty": len(_empty_primary_keys),
                                "unique_keys": len(_c),
                            },
                            "timestamp": __import__("time").time(),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion

    # #region agent log (session 5b5324) после пересборки primary plate_assignments
    if _opt_debug_enabled():
        try:
            _prim_sample = [
                {
                    "kp_id": p.get("kp_id"),
                    "plate_name": (p.get("plate_name") or "")[:50],
                    "identity_match_type": p.get("identity_match_type"),
                }
                for p in result["plate_assignments"][:5]
            ]
            with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                _f.write(
                    __import__("json").dumps(
                        {
                            "sessionId": "5b5324",
                            "hypothesisId": "H_primary_pa",
                            "location": "core/optimization/optimize_2d/finalize.py:primary_pa",
                            "message": "primary plate_assignments count and sample",
                            "data": {
                                "primary_plate_assignments_count": len(result["plate_assignments"]),
                                "sample": _prim_sample,
                            },
                            "timestamp": __import__("time").time(),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion

    print(f"[OPT_2D] ✓ Создано остатков: {len(result['rests_created'])}")
    # ========== КОНЕЦ НОВОЙ ЛОГИКИ ==========

    # Вторичные резы: получаем identity из того же общего slot-ledger.
    _secondary_attribution_log = []  # (target_key, kp_id, plate_name, match_type) для лога 5b5324
    _empty_secondary_keys = []
    for cut in result["secondary_cuts"]:
        target_key = cut.get("target_order_key")
        plate_info = _next_slot_info(slot_lists, slot_cursors, target_key) if target_key else {}
        if target_key and not plate_info:
            _empty_secondary_keys.append(target_key)
        if len(_secondary_attribution_log) < 50:
            _secondary_attribution_log.append(
                {
                    "target_key": list(target_key) if isinstance(target_key, tuple) else target_key,
                    "kp_id": plate_info.get("kp_id") if plate_info else None,
                    "plate_name": (plate_info.get("plate_name") or "")[:50] if plate_info else None,
                    "match_type": plate_info.get("identity_match_type") if plate_info else "empty",
                }
            )
        # #region agent log: secondary plate kp_id (H2, H5)
        if _opt_debug_enabled() and not plate_info and target_key:
            try:
                with _dbg_open_append(_DEBUG_LOG_COMMON) as _f:
                    _f.write(
                        __import__("json").dumps(
                            {
                                "hypothesisId": "H2",
                                "location": "core/optimization/optimize_2d/finalize.py:secondary_empty",
                                "message": "secondary plate_info empty",
                                "data": {
                                    "target_key": list(target_key)
                                    if isinstance(target_key, tuple)
                                    else target_key,
                                    "output_length": cut.get("lengths", [None])[0],
                                    "output_width": cut.get("cuts", [None])[0],
                                },
                                "timestamp": __import__("time").time(),
                            },
                            ensure_ascii=False,
                        )
                        + "\n"
                    )
            except Exception:
                pass
        # #endregion
        cut["kp_id"] = plate_info.get("kp_id") if plate_info else None
        cut["customer"] = plate_info.get("customer") if plate_info else None
        cut["kp_date"] = plate_info.get("kp_date") if plate_info else None
        cut["plate_name"] = plate_info.get("plate_name") if plate_info else None
        cut["identity_match_type"] = (
            plate_info.get("identity_match_type") if plate_info else "secondary_unmapped"
        )
        cut["concrete_grade"] = plate_info.get("concrete_grade") if plate_info else None

        result["plate_assignments"].append(
            {
                "length": cut["lengths"][0],
                "width": cut["cuts"][0],
                "source": "secondary",
                "source_rest": cut["source"],
                "kp_id": cut.get("kp_id"),
                "customer": cut.get("customer"),
                "kp_date": cut.get("kp_date"),
                "plate_name": cut.get("plate_name"),
                "load_code": cut.get("load_code", 800),
                "identity_match_type": cut.get("identity_match_type"),
                "concrete_grade": cut.get("concrete_grade"),
                "unit_id": cut.get("secondary_instance_id"),
                "parent_unit_id": cut.get("parent_instance_id"),
            }
        )
    # #region agent log (ef42ae: H1 оптимизатор дал вторичку без родителя; H3 «ошибочно вторичный» по watchlist)
    if _opt_debug_enabled():
        try:
            import json as _ef42_json
            import time as _ef42_time

            _sec_asg = [p for p in result["plate_assignments"] if p.get("source") == "secondary"]
            _null_par = [p for p in _sec_asg if not p.get("parent_unit_id")]
            _watch_subs = ("25,4-3", "63,9-5,3", "42,6-5,3", "25,4-3,0", "63,9-5,3-10")
            _watch_sec = [
                {
                    "unit_id": p.get("unit_id"),
                    "parent_unit_id": p.get("parent_unit_id"),
                    "length": p.get("length"),
                    "width": p.get("width"),
                    "plate_name": (p.get("plate_name") or "")[:160],
                    "identity_match_type": p.get("identity_match_type"),
                }
                for p in _sec_asg
                if any(s in str(p.get("plate_name") or "") for s in _watch_subs)
            ]
            _prim_asg = [p for p in result["plate_assignments"] if p.get("source") != "secondary"]
            _watch_prim = [
                {
                    "length": p.get("length"),
                    "width": p.get("width"),
                    "plate_name": (p.get("plate_name") or "")[:160],
                    "source": p.get("source"),
                }
                for p in _prim_asg
                if any(s in str(p.get("plate_name") or "") for s in _watch_subs)
            ]
            _raw_sec_no_parent = [
                {
                    "secondary_instance_id": c.get("secondary_instance_id"),
                    "parent_instance_id": c.get("parent_instance_id"),
                    "target_order_key": list(c["target_order_key"])
                    if isinstance(c.get("target_order_key"), tuple)
                    else c.get("target_order_key"),
                    "lengths": c.get("lengths"),
                    "cuts": c.get("cuts"),
                }
                for c in (result.get("secondary_cuts") or [])
                if not c.get("parent_instance_id")
            ][:120]
            with _dbg_open_append(_DEBUG_LOG_EF42AE) as _ef42_f:
                _ef42_f.write(
                    _ef42_json.dumps(
                        {
                            "sessionId": "ef42ae",
                            "hypothesisId": "H1_H3",
                            "location": "core/optimization/optimize_2d/finalize.py:ef42ae",
                            "message": "secondary assignments: parent null count, raw secondary_cuts without parent, SKU watchlist primary vs secondary",
                            "data": {
                                "n_secondary_assignments": len(_sec_asg),
                                "n_secondary_null_parent": len(_null_par),
                                "n_raw_secondary_cuts": len(result.get("secondary_cuts") or []),
                                "raw_secondary_no_parent_count": len(
                                    [
                                        c
                                        for c in (result.get("secondary_cuts") or [])
                                        if not c.get("parent_instance_id")
                                    ]
                                ),
                                "null_parent_assignments_head": [
                                    {
                                        "unit_id": p.get("unit_id"),
                                        "parent_unit_id": p.get("parent_unit_id"),
                                        "length": p.get("length"),
                                        "width": p.get("width"),
                                        "plate_name": (p.get("plate_name") or "")[:120],
                                    }
                                    for p in _null_par[:100]
                                ],
                                "raw_secondary_no_parent_head": _raw_sec_no_parent,
                                "watchlist_secondary": _watch_sec,
                                "watchlist_primary_rows": _watch_prim[:60],
                            },
                            "timestamp": int(_ef42_time.time() * 1000),
                        },
                        ensure_ascii=False,
                        default=str,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion
    # #region agent log (session 5b5324) после вторичных резов
    if _opt_debug_enabled():
        try:
            _sec_count = sum(1 for p in result["plate_assignments"] if p.get("source") == "secondary")
            with _dbg_open_append(_DEBUG_LOG_5b5324) as _f:
                _f.write(
                    __import__("json").dumps(
                        {
                            "sessionId": "5b5324",
                            "hypothesisId": "H_secondary",
                            "location": "core/optimization/optimize_2d/finalize.py:secondary",
                            "message": "secondary attributions and plate_assignments",
                            "data": {
                                "secondary_attribution_sample": _secondary_attribution_log[:40],
                                "plate_assignments_total": len(result["plate_assignments"]),
                                "secondary_count": _sec_count,
                            },
                            "timestamp": __import__("time").time(),
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )
        except Exception:
            pass
    # #endregion
    # #region agent log
    if _opt_debug_enabled():
        try:
            _slot_exhausted_by_key = {}
            if _empty_primary_keys:
                _slot_exhausted_by_key["primary"] = {
                    str(list(k)): int(v) for k, v in Counter(_empty_primary_keys).items()
                }
            if _empty_secondary_keys:
                _slot_exhausted_by_key["secondary"] = {
                    str(list(k)): int(v) for k, v in Counter(_empty_secondary_keys).items()
                }
            _slot_cursor_overview = []
            for _k in list(slot_lists.keys())[:80]:
                _slot_cursor_overview.append(
                    {
                        "key": list(_k),
                        "cursor": int(slot_cursors.get(_k, 0)),
                        "slots": int(len(slot_lists.get(_k, []))),
                    }
                )
            _debug_runtime_write_648532(
                "run1",
                "H2_slot_exhaustion",
                "core/optimization/optimize_2d/finalize.py:slot_exhaustion",
                "Slot consumption and exhausted attribution keys",
                {
                    "plate_assignments_total": int(len(result.get("plate_assignments", []))),
                    "empty_primary_count": int(len(_empty_primary_keys)),
                    "empty_secondary_count": int(len(_empty_secondary_keys)),
                    "slot_exhausted_by_key": _slot_exhausted_by_key,
                    "slot_cursor_overview": _slot_cursor_overview,
                },
            )
        except Exception:
            pass
    # #endregion

    print(f"[OPT_2D] OK! Готово! Использовано {result['total_plates']} плит")
    print(f"[OPT_2D] Создано {len(result['plate_assignments'])} готовых плит")
    print(f"[OPT_2D] Остатков использовано вторично: {len(result['rests_used'])}")
    # #region agent log: result counts (H5) + plates by key (H_rescue_trace)
    if _opt_debug_enabled():
        try:
            _dbg_open_append(_DEBUG_LOG_COMMON).write(
                __import__("json").dumps(
                    {
                        "sessionId": "debug-session",
                        "runId": "run1",
                        "hypothesisId": "H5",
                        "location": "core/optimization/optimize_2d/finalize.py:h5",
                        "message": "plate_assignments count",
                        "data": {
                            "len_plate_assignments": len(result["plate_assignments"]),
                            "total_plates": result.get("total_plates", 0),
                            "demand_sum": sum(demand_2d.values()),
                        },
                        "timestamp": __import__("time").time() * 1000,
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
        except Exception:
            pass
    if _opt_debug_enabled():
        try:
            _norm_lc = normalize_load_code
            _by_key = {}
            for cut in result.get("primary_cuts", []):
                lc = _norm_lc(cut.get("load_code", 8))
                for L in cut.get("lengths", []):
                    k = (round(float(L), 2), int(cut.get("width", 0)), lc)
                    _by_key[k] = _by_key.get(k, 0) + 1
            for cut in result.get("secondary_cuts", []):
                lengths = cut.get("lengths", [])
                widths = cut.get("cuts", [])
                tk = cut.get("target_order_key")
                lc = _norm_lc(tk[2] if isinstance(tk, (tuple, list)) and len(tk) > 2 else 8)
                L = float(lengths[0]) if lengths else 0
                W = (
                    int(widths[0])
                    if widths
                    else (int(tk[1]) if isinstance(tk, (tuple, list)) and len(tk) > 1 else 0)
                )
                if L and W:
                    k = (round(L, 2), W, lc)
                    _by_key[k] = _by_key.get(k, 0) + 1
            _dbg_open_append(_DEBUG_LOG_COMMON).write(
                __import__("json").dumps(
                    {
                        "hypothesisId": "H_opt_plates_by_key",
                        "location": "core/optimization/optimize_2d/finalize.py:by_key",
                        "message": "optimizer output plates by (length, width, load_code)",
                        "data": {
                            "plates_by_key": {str(list(k)): v for k, v in _by_key.items()},
                            "total": sum(_by_key.values()),
                        },
                        "timestamp": __import__("time").time() * 1000,
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
        except Exception:
            pass
    # #endregion
    return opt_ok(result, partial=(solver_status != "Optimal"))
