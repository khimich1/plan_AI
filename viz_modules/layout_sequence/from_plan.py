# -*- coding: utf-8 -*-
"""Построение последовательности из плана оптимизации (_build_sequence_from_plan)."""

from __future__ import annotations

import logging
import time
from collections import defaultdict
from typing import Any

import core.config_and_data as cfg

from core.optimization.layout_runtime_snapshot import LayoutSequenceCfgSlice

from viz_modules.layout_sequence.debug_trace import append_json_line, layout_sequence_trace_context
from viz_modules.layout_sequence.deps import LayoutSequenceDeps
from viz_modules.layout_sequence.helpers import (
    choose_best_separator,
    ensure_sequence_layout_uid,
    get_reinforcement_from_map,
    split_group_into_subgroups,
)
from viz_modules.layout_sequence.secondary_ops import (
    append_parent_variants_to_secondary_cuts_info,
    extract_orphan_secondaries_as_synthetic_primary_cuts,
    merge_atomic_secondaries_by_shared_parent,
    plan_stock_width_mm,
    secondary_geom_cut_key,
)


def _dispatch_agent_seq_debug(hypothesis_id: str, message: str, data: dict[str, Any]) -> None:
    """Через атрибут пакета: patch.object(viz_modules.layout_sequence, '_agent_seq_debug', ...) работает из тестов."""

    import viz_modules.layout_sequence as _ls_pkg

    _ls_pkg._agent_seq_debug(hypothesis_id, message, data)


def _build_sequence_from_plan(
    plan: dict[str, Any],
    plate_label_func: Any,
    reinforcement_map: dict | None = None,
    *,
    layout_cfg: LayoutSequenceCfgSlice | None = None,
    deps: LayoutSequenceDeps | None = None,
) -> list[dict[str, Any]]:
    """Строит последовательность плит из плана оптимизации (совместимо с историческими тестами)."""
    resolved_deps = deps or LayoutSequenceDeps.create()
    with layout_sequence_trace_context(resolved_deps.traces):
        return _build_sequence_from_plan_impl(
            plan,
            plate_label_func,
            reinforcement_map,
            layout_cfg=layout_cfg,
            deps=resolved_deps,
        )


def _build_sequence_from_plan_impl(
    plan: dict[str, Any],
    plate_label_func: Any,
    reinforcement_map: dict | None,
    *,
    layout_cfg: LayoutSequenceCfgSlice | None,
    deps: LayoutSequenceDeps,
) -> list[dict[str, Any]]:
    if reinforcement_map is None:
        reinforcement_map = {}
    if layout_cfg is None:
        layout_cfg = LayoutSequenceCfgSlice.from_config_module(cfg)
    pl = layout_cfg.plate_lists
    _norm = layout_cfg.normalize_load_code
    vis_log = deps.log

    hypothesis_log = deps.traces.debug_log

    sequence: list[dict[str, Any]] = []
    _missing_reinf_logged: set[tuple[Any, Any, Any]] = set()
    _log_int = logging.getLogger(__name__)

    def _warn_missing_reinforcement(length: Any, width_mm: Any, load_code_from_cut: Any) -> None:
        if load_code_from_cut is None:
            return
        key = (round(length, 3), width_mm, load_code_from_cut)
        if key in _missing_reinf_logged:
            return
        _missing_reinf_logged.add(key)
        _log_int.warning(
            "[ВИЗУАЛИЗАЦИЯ] Нет армирования в карте для (length=%s, width_mm=%s, load_code=%s)",
            length,
            width_mm,
            load_code_from_cut,
        )

    use_2d_data = "plate_assignments" in plan and plan["plate_assignments"]
    all_plates_with_lengths: list[dict[str, Any]] = []

    if not use_2d_data:
        vis_log.info("[VISUAL] ⚠️ 2D данных нет, используем приближение")
        for plates, width_mm in [
            (pl.plates_1_2, 1200),
            (pl.plates_1_08, 1080),
            (pl.plates_0_32, 320),
            (pl.plates_0_46, 460),
            (pl.plates_0_70, 700),
            (pl.plates_0_72, 720),
            (pl.plates_0_86, 860),
            (pl.plates_0_88, 880),
            (pl.plates_0_74, 740),
            (pl.plates_0_48, 480),
            (pl.plates_0_50, 500),
            (pl.plates_0_34, 340),
        ]:
            for length in plates:
                all_plates_with_lengths.append({"length": length, "width": width_mm})
        all_plates_with_lengths.sort(key=lambda x: (-x["width"], -x["length"]))

    transverse_cut_map: dict[Any, dict[str, Any]] = {}
    if plan.get("transverse_cuts"):
        for tcut in plan["transverse_cuts"]:
            key = (tcut["source_length"], tcut["source_width"])
            transverse_cut_map[key] = {"target_length": tcut["target_length"], "remainder": tcut["remainder"]}

    _dispatch_agent_seq_debug(
        "H0",
        "plan_input_after_maps",
        {
            "use_2d_data": use_2d_data,
            "n_primary_cuts": len(plan.get("primary_cuts") or []),
            "n_secondary_cuts_records": len(plan.get("secondary_cuts") or []),
            "n_transverse_cuts": len(plan.get("transverse_cuts") or []),
            "transverse_keys_sample": [list(k) for k in list(transverse_cut_map.keys())[:15]],
            "primary_preview": [
                {
                    "plate_name": (c.get("plate_name") or "")[:100],
                    "width": c.get("width"),
                    "rest": c.get("rest"),
                    "qty": c.get("qty"),
                    "primary_instance_id": c.get("primary_instance_id"),
                    "primary_instance_ids": [str(x) for x in (c.get("primary_instance_ids") or [])[:12]],
                    "lengths_head": (c.get("lengths") or [])[:12],
                    "load_code": c.get("load_code"),
                }
                for c in (plan.get("primary_cuts") or [])[:30]
            ],
            "secondary_preview": [
                {
                    "source_mm": sc.get("source"),
                    "qty": sc.get("qty"),
                    "cuts": sc.get("cuts"),
                    "source_lengths_head": (sc.get("source_lengths") or [])[:12],
                    "parent_instance_ids_head": [str(x) for x in (sc.get("parent_instance_ids") or [])[:12]],
                    "secondary_instance_ids_head": [str(x) for x in (sc.get("secondary_instance_ids") or [])[:12]],
                    "target_order_key": sc.get("target_order_key"),
                }
                for sc in (plan.get("secondary_cuts") or [])[:30]
            ],
        },
    )

    secondary_cuts_info: dict[Any, Any] = {}
    secondary_cuts_by_parent: dict[str, list[dict]] = defaultdict(list)
    secondary_total_from_plan = 0
    secondary_attached_total = 0
    unmatched_by_reason: defaultdict[str, int] = defaultdict(int)
    legacy_secondary_match_used = 0
    if plan.get("secondary_cuts"):
        for sec_cut in plan["secondary_cuts"]:
            source_mm = sec_cut["source"]
            pieces = sec_cut.get("pieces", 1)
            cuts_list = sec_cut.get("cuts", [])
            qty = sec_cut["qty"]
            secondary_total_from_plan += int(qty) * max(1, int(pieces))

            source_lengths_list = sec_cut.get("source_lengths", [])
            target_lengths_list = sec_cut.get("lengths", [])
            target_order_key = sec_cut.get("target_order_key")
            target_load_code = _norm(target_order_key[2]) if (target_order_key and len(target_order_key) > 2) else None
            parent_ids_list = sec_cut.get("parent_instance_ids") or []
            secondary_ids_list = sec_cut.get("secondary_instance_ids") or []

            pattern: list[dict[str, Any]] = []
            if cuts_list:
                target_width_mm = cuts_list[0]
                for _ in range(pieces):
                    pattern.append(
                        {
                            "width": target_width_mm / 1000.0,
                            "width_mm": target_width_mm,
                            "source_width_mm": source_mm,
                            "label": None,
                            "target_length": target_lengths_list[0] if target_lengths_list else None,
                            "target_load_code": target_load_code,
                        }
                    )

            for i in range(qty):
                source_length = source_lengths_list[i] if i < len(source_lengths_list) else 6.0
                key = secondary_geom_cut_key(source_length, source_mm)
                parent_instance_id = (
                    parent_ids_list[i] if i < len(parent_ids_list) else sec_cut.get("parent_instance_id")
                )
                secondary_instance_id = (
                    secondary_ids_list[i] if i < len(secondary_ids_list) else sec_cut.get("secondary_instance_id")
                )

                variant = {
                    "pattern": [segment.copy() for segment in pattern],
                    "qty": 1,
                    "used": 0,
                    "target_order_key": target_order_key,
                    "parent_instance_id": parent_instance_id,
                    "secondary_instance_id": secondary_instance_id,
                    "geom_key": key,
                }
                if parent_instance_id:
                    secondary_cuts_by_parent[str(parent_instance_id)].append(variant)
                else:
                    if key not in secondary_cuts_info:
                        secondary_cuts_info[key] = []
                    secondary_cuts_info[key].append(variant)
        merge_atomic_secondaries_by_shared_parent(secondary_cuts_by_parent=secondary_cuts_by_parent, logger=_log_int)
        append_parent_variants_to_secondary_cuts_info(
            secondary_cuts_info=secondary_cuts_info, secondary_cuts_by_parent=secondary_cuts_by_parent
        )

    if plan.get("secondary_cuts"):
        _geom_variant_counts = {f"{k[0]}_{k[1]}": len(v) for k, v in secondary_cuts_info.items()}
        _dispatch_agent_seq_debug(
            "H1",
            "secondary_index_after_phase_A",
            {
                "n_geom_keys": len(secondary_cuts_info),
                "n_parent_buckets": len(secondary_cuts_by_parent),
                "geom_key_variant_counts": _geom_variant_counts,
                "parent_bucket_keys_head": list(secondary_cuts_by_parent.keys())[:50],
                "variants_per_parent_top": sorted(
                    ((str(pid), len(lst)) for pid, lst in secondary_cuts_by_parent.items()),
                    key=lambda x: -x[1],
                )[:20],
            },
        )

    synthetic_orphan_primary_cuts: list[dict] = []
    if plan.get("secondary_cuts"):
        synthetic_orphan_primary_cuts = extract_orphan_secondaries_as_synthetic_primary_cuts(
            secondary_cuts_info=secondary_cuts_info,
            plan=plan,
            plate_width_mm=plan_stock_width_mm(plan),
            normalize_load_code=_norm,
        )

    all_primary_cuts = plan.get("primary_cuts", [])
    solid_cuts = [cut for cut in all_primary_cuts if cut["rest"] == 0]

    for cut in solid_cuts:
        length = cut["lengths"][0] if cut.get("lengths") else 6.0
        width_mm = cut["width"]
        lc = _norm(cut.get("load_code", 8))
        cut["reinforcement"] = (
            get_reinforcement_from_map(
                reinforcement_map, length, width_mm, lc, hypothesis_debug_log=hypothesis_log
            )
            or 999.0
        )
    solid_cuts.sort(key=lambda x: (x.get("reinforcement", 999.0), -x["lengths"][0] if x.get("lengths") else 0))

    cut_with_rest_raw = [cut for cut in all_primary_cuts if cut["rest"] > 0]
    if synthetic_orphan_primary_cuts:
        cut_with_rest_raw.extend(synthetic_orphan_primary_cuts)

    for cut in cut_with_rest_raw:
        length = cut["lengths"][0] if cut.get("lengths") else 6.0
        width_mm = cut["width"]
        lc = _norm(cut.get("load_code", 8))
        cut["reinforcement"] = (
            get_reinforcement_from_map(
                reinforcement_map, length, width_mm, lc, hypothesis_debug_log=hypothesis_log
            )
            or 999.0
        )

    cut_with_rest = sorted(
        cut_with_rest_raw,
        key=lambda x: (x.get("reinforcement", 999.0), x["width"], x["rest"]),
    )

    vis_log.info(f"[VISUAL] Разделение: {len(solid_cuts)} типов целых плит, {len(cut_with_rest)} типов с резом")
    if solid_cuts:
        vis_log.info(
            f"[VISUAL] Целые плиты (сортировка по армированию): "
            f"{[(c['width'], c['qty'], c.get('reinforcement', '?')) for c in solid_cuts[:5]]}"
        )

    from itertools import groupby

    cut_groups = [
        list(group)
        for _key, group in groupby(
            cut_with_rest, key=lambda x: (x["width"], x["rest"], x.get("reinforcement", 999.0))
        )
    ]

    if cut_groups:
        vis_log.info(f"[VISUAL] Найдено {len(cut_groups)} групп резов (сгруппировано по рез+армирование):")
        for i, group in enumerate(cut_groups, 1):
            vis_log.info(
                f"[VISUAL]   Группа {i}: width={group[0]['width']}мм, rest={group[0]['rest']}мм, "
                f"армирование={group[0].get('reinforcement', '?'):.1f}, "
                f"плит={sum(c['qty'] for c in group)}"
            )

    ordered_cuts: list[dict[str, Any]] = []

    solid_cuts_list: list[dict[str, Any]] = []
    for cut in solid_cuts:
        lengths = cut.get("lengths", [])
        for i in range(cut["qty"]):
            single_cut = cut.copy()
            single_cut["qty"] = 1
            single_cut["lengths"] = [lengths[i]] if i < len(lengths) else [lengths[0] if lengths else 6.0]
            solid_cuts_list.append(single_cut)

    vis_log.info(f"[VISUAL] Развёрнуто {len(solid_cuts_list)} отдельных целых плит для разделителей")

    if solid_cuts_list:
        first_plate = solid_cuts_list.pop(0)
        ordered_cuts.append(first_plate)
        first_width = first_plate.get("width", 1200)
        vis_log.info(f"[VISUAL] ✓ Первая плита: целая {first_width}мм")

    for i, cut_group in enumerate(cut_groups):
        total_group_length = sum(
            cut["lengths"][0] * cut["qty"] for cut in cut_group if cut.get("lengths")
        )
        vis_log.info(
            f"[VISUAL] Группа резов #{i+1}: width={cut_group[0]['width']}мм, "
            f"rest={cut_group[0]['rest']}мм, типов={len(cut_group)}, длина={total_group_length:.1f}м"
        )

        if total_group_length > 90.0:
            vis_log.info(
                f"[VISUAL] ⚠️ Группа #{i+1} слишком большая ({total_group_length:.1f}м > 90м), разбиваем на подгруппы"
            )
            subgroups = split_group_into_subgroups(cut_group, max_length=90.0, log=vis_log)
            for j, subgroup in enumerate(subgroups):
                ordered_cuts.extend(subgroup)
                vis_log.info(f"[VISUAL]   Добавлена подгруппа #{i+1}.{j+1}: {len(subgroup)} плит")
                if j < len(subgroups) - 1 and solid_cuts_list:
                    separator = solid_cuts_list.pop(0)
                    separator["is_separator"] = True
                    ordered_cuts.append(separator)
                    vis_log.info("[VISUAL]   ✓ Разделитель между подгруппами: целая плита")
        else:
            ordered_cuts.extend(cut_group)
            vis_log.info(f"[VISUAL] Добавлена группа резов #{i+1} (влезает в дорожку)")

        if i < len(cut_groups) - 1 and solid_cuts_list:
            next_group = cut_groups[i + 1]
            best_idx = choose_best_separator(solid_cuts_list, next_group, reinforcement_map, log=vis_log)
            if best_idx is not None:
                separator = solid_cuts_list.pop(best_idx)
                separator["is_separator"] = True
                ordered_cuts.append(separator)
                vis_log.info("[VISUAL] ✓ Разделитель (is_separator=True): целая плита между группами")
            else:
                if solid_cuts_list:
                    fallback_sep = solid_cuts_list.pop(0)
                    fallback_sep["is_separator"] = True
                    ordered_cuts.append(fallback_sep)
                    vis_log.info("[VISUAL] ✓ Разделитель: целая плита между группами (fallback, is_separator=True)")

    if solid_cuts_list:
        ordered_cuts.extend(solid_cuts_list)
        vis_log.info(f"[VISUAL] Добавлено {len(solid_cuts_list)} оставшихся целых плит в конец")

    for oi, ocut in enumerate(ordered_cuts):
        if ocut.get("rest", 0) <= 0:
            continue
        _pids = ocut.get("primary_instance_ids") or []
        if ocut.get("primary_instance_id") and not _pids:
            _pids = [ocut.get("primary_instance_id")]
        _dispatch_agent_seq_debug(
            "H2",
            "ordered_cut_split_row_phase_B",
            {
                "ordered_idx": oi,
                "plate_name": (ocut.get("plate_name") or "")[:120],
                "width_mm": ocut.get("width"),
                "rest_mm": ocut.get("rest"),
                "qty": ocut.get("qty"),
                "primary_instance_ids": [str(x) for x in _pids],
                "len_primary_ids_vs_qty": {"len_pids": len(_pids), "qty": ocut.get("qty")},
                "lengths": (ocut.get("lengths") or []),
                "load_code": ocut.get("load_code"),
            },
        )

    for cut in ordered_cuts:
        width_mm = cut["width"]
        rest_mm = cut["rest"]
        qty = cut["qty"]
        primary_instance_ids = cut.get("primary_instance_ids") or []
        if cut.get("primary_instance_id") and not primary_instance_ids:
            primary_instance_ids = [cut.get("primary_instance_id")]

        if use_2d_data and "lengths" in cut:
            lengths_for_cut = cut["lengths"]
        else:
            matching_plates = [p for p in all_plates_with_lengths if p["width"] == width_mm]
            lengths_for_cut = [p["length"] for p in matching_plates[:qty]]
            while len(lengths_for_cut) < qty:
                lengths_for_cut.append(6.0 if not matching_plates else matching_plates[0]["length"])

        for i in range(qty):
            length = lengths_for_cut[i] if i < len(lengths_for_cut) else 6.0
            parent_instance_id = (
                primary_instance_ids[i] if i < len(primary_instance_ids) else cut.get("primary_instance_id")
            )
            plate_uids = cut.get("plate_uids") or []
            plate_uid = plate_uids[i] if i < len(plate_uids) else cut.get("plate_uid")

            kp_id = cut.get("kp_id")
            customer = cut.get("customer")
            kp_date = cut.get("kp_date")
            plate_name_from_cut = cut.get("plate_name")

            transverse_cut_info = transverse_cut_map.get((length, width_mm))
            if cut.get("ignore_transverse"):
                transverse_cut_info = None

            if transverse_cut_info:
                _dispatch_agent_seq_debug(
                    "H4",
                    "phase_C_transverse_branch_skips_split_secondary",
                    {
                        "plate_name": (plate_name_from_cut or "")[:120],
                        "length_m": length,
                        "width_mm": width_mm,
                        "rest_mm": rest_mm,
                        "unit_idx_in_cut": i,
                        "qty": qty,
                        "parent_instance_id": str(parent_instance_id) if parent_instance_id else None,
                        "transverse_target_length": transverse_cut_info.get("target_length"),
                        "transverse_remainder": transverse_cut_info.get("remainder"),
                    },
                )
                width_m = width_mm / 1000.0
                load_code_from_cut = _norm(cut.get("load_code", 8))
                reinforcement = get_reinforcement_from_map(
                    reinforcement_map, length, width_mm, load_code_from_cut, hypothesis_debug_log=hypothesis_log
                )
                if reinforcement is None:
                    _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                sequence.append(
                    {
                        "length": length,
                        "mode": "transverse",
                        "target_length": transverse_cut_info["target_length"],
                        "remainder": transverse_cut_info["remainder"],
                        "width": width_m,
                        "load_code": load_code_from_cut,
                        "label_target": plate_label_func(transverse_cut_info["target_length"], width_m, load_code_from_cut),
                        "label_remainder": f'Остаток {transverse_cut_info["remainder"]:.2f}м'.replace(".", ",")
                        if transverse_cut_info["remainder"] > 0.1
                        else "",
                        "reinforcement": reinforcement,
                        "kp_id": kp_id,
                        "customer": customer,
                        "kp_date": kp_date,
                        "plate_name": plate_name_from_cut,
                        "unit_id": parent_instance_id,
                        "layout_uid": str(parent_instance_id)
                        if parent_instance_id
                        else f"transverse:{len(sequence)}",
                    }
                )
            else:
                main_w = width_mm / 1000.0
                rest_w = rest_mm / 1000.0
                fake_rest_override = False

                if width_mm == 1080 and rest_mm == 0:
                    rest_mm = 120
                    rest_w = 0.12
                    fake_rest_override = True

                if rest_mm == 0:
                    load_code_from_cut = _norm(cut.get("load_code", 8))
                    reinforcement = get_reinforcement_from_map(
                        reinforcement_map, length, width_mm, load_code_from_cut, hypothesis_debug_log=hypothesis_log
                    )
                    if reinforcement is None:
                        _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                    if abs(length - 7.1) < 0.05 and load_code_from_cut in (12, 12.5, 13):
                        try:
                            append_json_line(
                                hypothesis_log,
                                {
                                    "hypothesisId": "H71item2d",
                                    "location": "layout_sequence:sequence_append_solid_2d",
                                    "message": "71-12п item в sequence (2d)",
                                    "data": {
                                        "length": length,
                                        "width_mm": width_mm,
                                        "load_code_from_cut": load_code_from_cut,
                                        "reinforcement": reinforcement,
                                    },
                                    "timestamp": __import__("time").time() * 1000,
                                },
                            )
                        except Exception:
                            pass
                    if not reinforcement:
                        vis_log.info(
                            f"[VISUAL] ⚠️ Армирование не найдено для целой плиты: "
                            f"{length}м x {width_mm}мм (load_code={load_code_from_cut})"
                        )
                        vis_log.info(f"[VISUAL]    Доступные ключи в карте: {list(reinforcement_map.keys())[:5]}")
                    else:
                        vis_log.info(
                            f"[VISUAL] ✓ Армирование найдено для целой плиты: "
                            f"{length}м x {width_mm}мм = {reinforcement:.1f}"
                        )
                    is_separator = cut.get("is_separator", False)
                    sequence.append(
                        {
                            "length": length,
                            "mode": "solid",
                            "width": main_w,
                            "load_code": load_code_from_cut,
                            "label": plate_label_func(length, main_w, load_code_from_cut),
                            "reinforcement": reinforcement,
                            "is_separator": is_separator,
                            "kp_id": kp_id,
                            "customer": customer,
                            "kp_date": kp_date,
                            "plate_name": plate_name_from_cut,
                            "plate_uid": plate_uid,
                            "unit_id": parent_instance_id,
                            "layout_uid": str(parent_instance_id)
                            if parent_instance_id
                            else f"solid:{len(sequence)}",
                        }
                    )
                else:
                    load_code_from_cut = _norm(cut.get("load_code", 8))
                    secondary_cuts_for_plate = None
                    chosen_variant = None
                    _matched_via = None
                    _geom_key_attach = secondary_geom_cut_key(length, rest_mm)
                    _pool_parent: list[Any] = []
                    _pool_geom: list[Any] = []
                    _pool_geom_for_log: list[Any] = []
                    if cut.get("skip_secondary_attachment"):
                        _matched_via = "emulated_primary"
                        secondary_attached_total += int(cut.get("emulated_secondary_units") or 1)
                    else:
                        if parent_instance_id:
                            _pool_parent = secondary_cuts_by_parent.get(str(parent_instance_id)) or []
                        for variant in _pool_parent:
                            if variant["used"] < variant["qty"] and (variant.get("pattern") or []):
                                chosen_variant = variant
                                _matched_via = "parent"
                                break
                        if not chosen_variant:
                            _pool_geom = secondary_cuts_info.get(_geom_key_attach) or []
                            if _pool_geom:
                                legacy_secondary_match_used += 1
                                for variant in _pool_geom:
                                    _variant_target_key = variant.get("target_order_key")
                                    _variant_target_lc = (
                                        _norm(_variant_target_key[2], default=8)
                                        if _variant_target_key and len(_variant_target_key) > 2
                                        else load_code_from_cut
                                    )
                                    if _variant_target_lc > load_code_from_cut:
                                        continue
                                    if variant["used"] < variant["qty"] and (variant.get("pattern") or []):
                                        chosen_variant = variant
                                        _matched_via = "geom"
                                        break
                        _pool_geom_for_log = secondary_cuts_info.get(_geom_key_attach) or []

                    if chosen_variant:
                        secondary_cuts_for_plate = []
                        _sec_tok_plan = chosen_variant.get("target_order_key")
                        for sec_idx, sec_cut_template in enumerate(chosen_variant["pattern"]):
                            _sec_ids_list = chosen_variant.get("secondary_instance_ids") or []
                            if sec_idx < len(_sec_ids_list) and _sec_ids_list[sec_idx]:
                                sec_unit_id = _sec_ids_list[sec_idx]
                            else:
                                sec_unit_id = chosen_variant.get("secondary_instance_id")
                            if sec_unit_id is None:
                                sec_unit_id = (
                                    f"{parent_instance_id or 'secondary'}:{length}:{rest_mm}:"
                                    f"{sec_idx}:{chosen_variant['used']}"
                                )
                            sec_width = sec_cut_template["width"]
                            sec_width_mm = sec_cut_template["width_mm"]
                            lc_l = sec_cut_template.get("target_load_code")
                            lc_use = _norm(lc_l) if lc_l is not None else load_code_from_cut

                            sec_transverse = transverse_cut_map.get((length, sec_width_mm))

                            if sec_transverse:
                                secondary_cuts_for_plate.append(
                                    {
                                        "width": sec_width,
                                        "label": f'[2] {plate_label_func(sec_transverse["target_length"], sec_width, lc_use)}',
                                        "transverse_cut": True,
                                        "target_length": sec_transverse["target_length"],
                                        "remainder": sec_transverse["remainder"],
                                        "load_code": lc_use,
                                        "target_order_key": _sec_tok_plan,
                                        "parent_unit_id": parent_instance_id,
                                        "unit_id": sec_unit_id,
                                    }
                                )
                            else:
                                target_length_tpl = sec_cut_template.get("target_length")
                                if target_length_tpl:
                                    secondary_cuts_for_plate.append(
                                        {
                                            "width": sec_width,
                                            "label": f"О {plate_label_func(target_length_tpl, sec_width, lc_use)}",
                                            "has_transverse": True,
                                            "target_length": target_length_tpl,
                                            "load_code": lc_use,
                                            "target_order_key": _sec_tok_plan,
                                            "parent_unit_id": parent_instance_id,
                                            "unit_id": sec_unit_id,
                                        }
                                    )
                                else:
                                    result_width = sec_cut_template["width"]
                                    source_width = sec_cut_template.get("source_width_mm", result_width * 1000) / 1000.0
                                    label_text = plate_label_func(length, result_width, lc_use)
                                    if abs(result_width - source_width) > 1e-6:
                                        label_text = f"О {label_text}"
                                    secondary_cuts_for_plate.append(
                                        {
                                            "width": result_width,
                                            "label": label_text,
                                            "load_code": lc_use,
                                            "target_order_key": _sec_tok_plan,
                                            "parent_unit_id": parent_instance_id,
                                            "unit_id": sec_unit_id,
                                        }
                                    )
                        chosen_variant["used"] += 1
                        secondary_attached_total += len(secondary_cuts_for_plate)
                    elif rest_mm > 0 and not cut.get("skip_secondary_attachment"):
                        if parent_instance_id:
                            unmatched_by_reason["parent_instance_id_not_found"] += 1
                        else:
                            unmatched_by_reason["key_not_found"] += 1

                    def _pool_free_n(pool: list) -> int:
                        return sum(
                            1
                            for v in pool
                            if v.get("used", 0) < v.get("qty", 0) and (v.get("pattern") or [])
                        )

                    _dispatch_agent_seq_debug(
                        "H3",
                        "phase_C_split_secondary_attach_attempt",
                        {
                            "plate_name": (plate_name_from_cut or "")[:120],
                            "length_m": length,
                            "width_mm": width_mm,
                            "rest_mm": rest_mm,
                            "geom_key": list(_geom_key_attach),
                            "unit_idx_in_cut": i,
                            "qty": qty,
                            "parent_instance_id": str(parent_instance_id) if parent_instance_id else None,
                            "n_pool_parent": len(_pool_parent),
                            "n_pool_parent_free": _pool_free_n(_pool_parent),
                            "n_pool_geom": len(_pool_geom_for_log),
                            "n_pool_geom_free": _pool_free_n(_pool_geom_for_log),
                            "matched_via": _matched_via,
                            "attached_pattern_segments": len(chosen_variant["pattern"])
                            if chosen_variant
                            else 0,
                            "chosen_secondary_id": (
                                str(
                                    chosen_variant.get("secondary_instance_ids")
                                    or chosen_variant.get("secondary_instance_id")
                                )
                                if chosen_variant
                                else None
                            ),
                            "chosen_variant_parent_in_plan": str(chosen_variant.get("parent_instance_id"))
                            if chosen_variant
                            else None,
                            "unmatched_increment": (
                                None
                                if cut.get("skip_secondary_attachment")
                                else (
                                    "parent_instance_id_not_found"
                                    if (not chosen_variant and rest_mm > 0 and parent_instance_id)
                                    else (
                                        "key_not_found"
                                        if (not chosen_variant and rest_mm > 0 and not parent_instance_id)
                                        else None
                                    )
                                )
                            ),
                        },
                    )

                    reinforcement = get_reinforcement_from_map(
                        reinforcement_map, length, width_mm, load_code_from_cut, hypothesis_debug_log=hypothesis_log
                    )
                    if reinforcement is None:
                        _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                    if reinforcement:
                        vis_log.info(
                            f"[VISUAL] ✓ Армирование найдено для плиты с резом: "
                            f"{length}м x {width_mm}мм = {reinforcement:.1f}"
                        )
                    sequence.append(
                        {
                            "length": length,
                            "mode": "split",
                            "main_w": main_w,
                            "rest_w": rest_w,
                            "load_code": load_code_from_cut,
                            "label_main": plate_label_func(length, main_w, load_code_from_cut),
                            "label_rest": (
                                "+0,12"
                                if fake_rest_override
                                else (f"+{rest_w:.2f}".replace(".", ",") if not secondary_cuts_for_plate else None)
                            ),
                            "secondary_cuts": secondary_cuts_for_plate,
                            "reinforcement": reinforcement,
                            "kp_id": kp_id,
                            "customer": customer,
                            "kp_date": kp_date,
                            "plate_name": plate_name_from_cut,
                            "plate_uid": plate_uid,
                            "unit_id": parent_instance_id,
                            "layout_uid": str(parent_instance_id)
                            if parent_instance_id
                            else f"split:{len(sequence)}",
                        }
                    )

    ensure_sequence_layout_uid(sequence, prefix="built")
    secondary_unmatched_total = max(0, secondary_total_from_plan - secondary_attached_total)

    _dispatch_agent_seq_debug(
        "H5",
        "phase_end_summary",
        {
            "secondary_total_from_plan": secondary_total_from_plan,
            "secondary_attached_total": secondary_attached_total,
            "secondary_unmatched_total": secondary_unmatched_total,
            "unmatched_by_reason": dict(unmatched_by_reason),
            "legacy_secondary_match_used": legacy_secondary_match_used,
            "sequence_len": len(sequence),
            "n_split_in_sequence": sum(1 for s in sequence if s.get("mode") == "split"),
            "n_transverse_in_sequence": sum(1 for s in sequence if s.get("mode") == "transverse"),
        },
    )

    if secondary_unmatched_total or legacy_secondary_match_used:
        _log_int.warning(
            "[LAYOUT_SEQUENCE] secondary mapping report: total=%s attached=%s unmatched=%s reasons=%s legacy_match_used=%s",
            secondary_total_from_plan,
            secondary_attached_total,
            secondary_unmatched_total,
            dict(unmatched_by_reason),
            legacy_secondary_match_used,
        )

    try:
        if secondary_unmatched_total or dict(unmatched_by_reason):
            _ef42_payload = {
                "sessionId": "ef42ae",
                "hypothesisId": "H2",
                "location": "viz_modules/layout_sequence/from_plan:_build_sequence_from_plan_impl:end",
                "message": "secondary plan vs attached to sequence (unmatched => sec unit_ids may miss tracks)",
                "data": {
                    "secondary_total_from_plan": secondary_total_from_plan,
                    "secondary_attached_total": secondary_attached_total,
                    "secondary_unmatched_total": secondary_unmatched_total,
                    "unmatched_by_reason": dict(unmatched_by_reason),
                    "legacy_secondary_match_used": legacy_secondary_match_used,
                },
                "timestamp": int(time.time() * 1000),
            }
            append_json_line(deps.traces.debug_ef42ae, _ef42_payload, ensure_ascii=False)
    except OSError:
        pass

    return sequence
