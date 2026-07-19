# -*- coding: utf-8 -*-
"""Phase C: PlateAudit checkpoints, post-correction, coverage verify, slot attribution."""

from __future__ import annotations

from collections import Counter

from core.domain.plate_order import normalize_load_code
from core.optimization.coverage_verify import verify_coverage
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

    print(f"[OPT_2D] OK! Готово! Использовано {result['total_plates']} плит")
    print(f"[OPT_2D] Создано {len(result['plate_assignments'])} готовых плит")
    print(f"[OPT_2D] Остатков использовано вторично: {len(result['rests_used'])}")
    return opt_ok(result, partial=(solver_status != "Optimal"))
