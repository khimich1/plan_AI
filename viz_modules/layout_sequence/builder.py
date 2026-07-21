# -*- coding: utf-8 -*-
"""Оркестрация build_layout_sequence: план, группы нагрузок, fallback-ветки."""
from __future__ import annotations

import logging
import time
from itertools import groupby
from pathlib import Path
from typing import TYPE_CHECKING, Any

from viz_modules.layout_sequence.debug_trace import LayoutSequenceTracePaths, append_json_line
from viz_modules.layout_sequence.deps import LayoutSequenceDeps
from viz_modules.layout_sequence.from_plan import _build_sequence_from_plan
from viz_modules.layout_sequence.helpers import (
    choose_best_separator,
    ensure_sequence_layout_uid,
    get_reinforcement_from_map,
    split_group_into_subgroups,
)
from viz_modules.layout_sequence.reinforcement_data import build_reinforcement_map_for_sequence
from viz_modules.layout_sequence.secondary_ops import secondary_geom_cut_key

if TYPE_CHECKING:
    from core.optimization.layout_runtime_snapshot import LayoutRuntimeSnapshot


def build_layout_sequence(
    *,
    runtime: LayoutRuntimeSnapshot | None = None,
    pb_db_path: Path | str | None = None,
    log: logging.Logger | None = None,
    traces: LayoutSequenceTracePaths | None = None,
) -> Any:
    """Формирует последовательность сегментов вдоль дорожки, РАЗДЕЛЁННУЮ ПО НАГРУЗКАМ.

    :param runtime: снимок замороженного плана OPT + срез cfg для раскладки (DIP-003). Если ``None``,
        собирается через ``build_layout_runtime_snapshot()`` — прежнее поведение.
    :param pb_db_path: путь к pb.db (по умолчанию из настроек приложения).
    :param log: логгер для сообщений ``[VISUAL]``; по умолчанию ``viz_modules.layout_sequence``.
    :param traces: пути NDJSON-логов агента; по умолчанию через ``get_debug_log_path``.
    """
    from core.optimization.layout_runtime_snapshot import build_layout_runtime_snapshot
    from core.reinforcement_db import get_reinforcement

    deps = LayoutSequenceDeps.create(pb_db_path=pb_db_path, log=log, traces=traces)

    if runtime is None:
        runtime = build_layout_runtime_snapshot()
    oh = runtime.opt_snapshot
    ls = runtime.layout_cfg
    pl = ls.plate_lists
    OPT_CASCADING_PLAN_BY_LOAD = oh.opt_cascading_plan_by_load
    OPT_CASCADING_PLAN = oh.opt_cascading_plan
    OPT_PLAN = oh.opt_plan
    OPT_WIDTH_PRIORITY = oh.opt_width_priority

    db_path = deps.pb_db_path
    reinforcement_map = build_reinforcement_map_for_sequence(
        plate_load_details=ls.plate_load_details,
        OPT_CASCADING_PLAN_BY_LOAD=OPT_CASCADING_PLAN_BY_LOAD,
        OPT_CASCADING_PLAN=OPT_CASCADING_PLAN,
        get_reinforcement=get_reinforcement,
        db_path=db_path,
        normalize_load_code=ls.normalize_load_code,
        deps=deps,
    )
    sequence: list[dict[str, Any]] = []

    def plate_label(L: float, W: float, load_code: int | None = None) -> str:
        """Метка плиты с правильной нагрузкой."""
        resolved_load = load_code
        if resolved_load is None and ls.plate_load_details:
            for key, _qty in ls.plate_load_details.items():
                plate_L, plate_W, plate_load = key[0], key[1], key[2]
                if abs(plate_L - L) < 0.05 and abs(plate_W - W) < 0.01:
                    resolved_load = plate_load
                    break
        if resolved_load is None:
            resolved_load = ls.get_load_code_for_plate(L, W, default=(6 if W < 1.0 else 8))
        return ls.make_plate_name(L, W, load_code=resolved_load)

    vis = deps.log
    traces_p = deps.traces

    vis.info(f"[VISUAL] Проверяем OPT_CASCADING_PLAN_BY_LOAD: {bool(OPT_CASCADING_PLAN_BY_LOAD)}")
    if OPT_CASCADING_PLAN_BY_LOAD:
        vis.info(
            f"[VISUAL] ✅ Используем группировку по нагрузкам! Групп: {len(OPT_CASCADING_PLAN_BY_LOAD)}"
        )
        try:
            _n_598665 = 0
            for _lg in sorted(OPT_CASCADING_PLAN_BY_LOAD.keys()):
                _p = OPT_CASCADING_PLAN_BY_LOAD[_lg]
                for _c in _p.get("primary_cuts", []):
                    _L = round(float((_c.get("lengths") or [6.0])[0]), 2)
                    _w = _c.get("width") or 1200
                    if abs(_L - 5.98) < 0.02 and _w == 665:
                        _n_598665 += _c.get("qty", 1)
            append_json_line(
                traces_p.debug_95694e,
                {
                    "sessionId": "95694e",
                    "hypothesisId": "H_95694e_plan_598665",
                    "location": "layout_sequence:build_layout_sequence:plans_in",
                    "message": "count 5.98/665 in primary_cuts across all load plans",
                    "data": {"count_598_665": _n_598665},
                    "timestamp": time.time(),
                },
                ensure_ascii=False,
            )
        except Exception:
            pass

        try:
            _n_508320 = _n_598530 = 0
            for _lg in sorted(OPT_CASCADING_PLAN_BY_LOAD.keys()):
                _p = OPT_CASCADING_PLAN_BY_LOAD[_lg]
                for _c in _p.get("primary_cuts", []):
                    _L = round(float((_c.get("lengths") or [6.0])[0]), 2)
                    _w = _c.get("width") or 1200
                    if abs(_L - 5.08) < 0.02 and _w == 320:
                        _n_508320 += _c.get("qty", 1)
                    if abs(_L - 5.98) < 0.02 and _w == 530:
                        _n_598530 += _c.get("qty", 1)
            append_json_line(
                traces_p.debug_95694e,
                {
                    "sessionId": "95694e",
                    "hypothesisId": "H_95694e_plan_rescue",
                    "location": "layout_sequence:build_layout_sequence:plans_in",
                    "message": "count 5.08/320 and 5.98/530 in primary_cuts across all load plans",
                    "data": {"count_508_320": _n_508320, "count_598_530": _n_598530},
                    "timestamp": time.time(),
                },
                ensure_ascii=False,
            )
        except Exception:
            pass

        all_sequences: list[dict[str, Any]] = []
        for load_group in sorted(OPT_CASCADING_PLAN_BY_LOAD.keys()):
            plan = OPT_CASCADING_PLAN_BY_LOAD[load_group]
            original_loads = plan.get("original_loads", [load_group])
            load_display_list = [ls.format_reinforcement_from_load_code(lc) for lc in original_loads]
            if len(load_display_list) > 1:
                load_display = ", ".join(load_display_list)
                label = f"Нагрузка {load_display}"
            else:
                label = f"Нагрузка {load_display_list[0]}"

            vis.info(f"[VISUAL] Обрабатываем группу {load_group} ({label})...")
            group_sequence = _build_sequence_from_plan(
                plan, plate_label, reinforcement_map, layout_cfg=ls, deps=deps
            )
            all_sequences.append(
                {
                    "load_code": load_group,
                    "original_loads": original_loads,
                    "sequence": group_sequence,
                    "label": label,
                }
            )
            vis.info(f"[VISUAL]   → {len(group_sequence)} плит в группе")

        try:
            _n_out = 0
            for _gr in all_sequences:
                for _it in _gr.get("sequence", []) or []:
                    _L = round(float(_it.get("length", 0) or _it.get("target_length", 0)), 2)
                    _w = _it.get("width") or _it.get("main_w") or 1.2
                    _w_mm = round(float(_w) * 1000) if float(_w) < 20 else round(float(_w))
                    if abs(_L - 5.98) < 0.02 and _w_mm == 665:
                        _n_out += 1
            append_json_line(
                traces_p.debug_95694e,
                {
                    "sessionId": "95694e",
                    "hypothesisId": "H_95694e_layout_out_598665",
                    "location": "layout_sequence:build_layout_sequence:plans_out",
                    "message": "count 5.98/665 in sequences after _build_sequence_from_plan",
                    "data": {"count_598_665": _n_out},
                    "timestamp": time.time(),
                },
                ensure_ascii=False,
            )
        except Exception:
            pass

        try:
            _n508, _n598 = 0, 0
            for _gr in all_sequences:
                for _it in _gr.get("sequence", []) or []:
                    _L = round(float(_it.get("length", 0) or _it.get("target_length", 0)), 2)
                    _w = _it.get("width") or _it.get("main_w") or 1.2
                    _w_mm = round(float(_w) * 1000) if float(_w) < 20 else round(float(_w))
                    if abs(_L - 5.08) < 0.02 and _w_mm == 320:
                        _n508 += 1
                    if abs(_L - 5.98) < 0.02 and _w_mm == 530:
                        _n598 += 1
            append_json_line(
                traces_p.debug_95694e,
                {
                    "sessionId": "95694e",
                    "hypothesisId": "H_95694e_layout_out_rescue",
                    "location": "layout_sequence:build_layout_sequence:plans_out",
                    "message": "count 5.08/320 and 5.98/530 in sequences after _build_sequence_from_plan",
                    "data": {"count_508_320": _n508, "count_598_530": _n598},
                    "timestamp": time.time(),
                },
                ensure_ascii=False,
            )
        except Exception:
            pass

        try:
            _target_keys = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
            _seq_by_key = {tuple(tk): 0 for tk in _target_keys}
            _total_in_sequence = 0
            for _gr in all_sequences:
                for s in _gr.get("sequence", []) or []:
                    _total_in_sequence += 1
                    L = round(float(s.get("length", 0) or s.get("target_length", 0)), 2)
                    w = s.get("width") or s.get("main_w") or 1.2
                    w_mm = round(float(w) * 1000) if float(w) < 20 else round(float(w))
                    lc = s.get("load_code", 8)
                    try:
                        lc = int(lc) if lc is not None else 8
                    except (TypeError, ValueError):
                        lc = 8
                    for tk in _target_keys:
                        if abs(L - tk[0]) <= 0.02 and w_mm == tk[1] and lc == tk[2]:
                            _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                            break
                    for sec in s.get("secondary_cuts", []):
                        sw = sec.get("width", 0)
                        sw_mm = round(float(sw) * 1000) if float(sw) < 20 else round(float(sw))
                        sl = round(float(sec.get("target_length") or L), 2)
                        for tk in _target_keys:
                            if abs(sl - tk[0]) <= 0.02 and sw_mm == tk[1]:
                                _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                                break
            _prim_total = sum(len(p.get("primary_cuts", [])) for p in OPT_CASCADING_PLAN_BY_LOAD.values())
            _seq_by_key_ser = [list(k) + [v] for k, v in _seq_by_key.items()]
            append_json_line(
                traces_p.debug_2d5c43,
                {
                    "sessionId": "2d5c43",
                    "hypothesisId": "H3",
                    "location": "layout_sequence:build_layout_sequence:grouped_return",
                    "message": "grouped path: sequence total and by key vs primary_cuts total",
                    "data": {
                        "total_from_primary": _prim_total,
                        "total_in_sequence": _total_in_sequence,
                        "sequence_by_key": _seq_by_key_ser,
                    },
                    "timestamp": time.time(),
                },
                ensure_ascii=False,
            )
        except Exception:
            pass

        return all_sequences

    vis.info(f"[VISUAL] Проверяем OPT_CASCADING_PLAN: {OPT_CASCADING_PLAN is not None}")
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get("primary_cuts"):
        sequence = _build_sequence_from_plan(
            OPT_CASCADING_PLAN, plate_label, reinforcement_map, layout_cfg=ls, deps=deps
        )
        ensure_sequence_layout_uid(sequence, prefix="single")
        return sequence

    # DEPRECATED / UNREACHABLE: дублирующий legacy-блок ниже никогда не выполняется —
    # ветка выше уже возвращает sequence через _build_sequence_from_plan (см. layout_reinforcement_order).
    if OPT_CASCADING_PLAN and OPT_CASCADING_PLAN.get("primary_cuts"):
        vis.info("[VISUAL] OK: Используем каскадную оптимизацию для визуализации")
        vis.info(f"[VISUAL] Первичных резов: {len(OPT_CASCADING_PLAN.get('primary_cuts', []))}")
        vis.info(f"[VISUAL] Вторичных резов: {len(OPT_CASCADING_PLAN.get('secondary_cuts', []))}")

        use_2d_data = "plate_assignments" in OPT_CASCADING_PLAN and OPT_CASCADING_PLAN["plate_assignments"]

        all_plates_with_lengths: list[dict[str, Any]] = []
        if use_2d_data:
            vis.info("[VISUAL] ТОЧНО: Используем 2D данные с точными длинами")
        else:
            vis.info("[VISUAL] ВНИМАНИЕ: 2D данных нет, используем приближение")
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

        transverse_cut_map: dict[Any, Any] = {}
        if OPT_CASCADING_PLAN.get("transverse_cuts"):
            for tcut in OPT_CASCADING_PLAN["transverse_cuts"]:
                key = (tcut["source_length"], tcut["source_width"])
                transverse_cut_map[key] = {"target_length": tcut["target_length"], "remainder": tcut["remainder"]}
        vis.info(
            f"[VISUAL] Найдено {len(transverse_cut_map)} типов поперечных резов: {list(transverse_cut_map.keys())}"
        )

        secondary_cuts_info: dict[Any, Any] = {}
        if OPT_CASCADING_PLAN.get("secondary_cuts"):
            for sec_cut in OPT_CASCADING_PLAN["secondary_cuts"]:
                source_mm = sec_cut["source"]
                pieces = sec_cut.get("pieces", 1)
                cuts_list = sec_cut.get("cuts", [])
                qty = sec_cut["qty"]

                source_lengths_list = sec_cut.get("source_lengths", [])
                target_lengths_list = sec_cut.get("lengths", [])
                target_order_key = sec_cut.get("target_order_key")
                target_load_code = (
                    ls.normalize_load_code(target_order_key[2])
                    if (target_order_key and len(target_order_key) > 2)
                    else None
                )

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

                    if key not in secondary_cuts_info:
                        secondary_cuts_info[key] = []

                    secondary_cuts_info[key].append(
                        {
                            "pattern": [segment.copy() for segment in pattern],
                            "qty": 1,
                            "used": 0,
                            "target_order_key": target_order_key,
                        }
                    )

        vis.info(f"[VISUAL] Создано {len(secondary_cuts_info)} вариантов вторичных резов:")
        for (src_len, src_w), variants in secondary_cuts_info.items():
            for idx, info in enumerate(variants, start=1):
                pattern_desc = ", ".join([f"{c['width_mm']}мм" for c in info["pattern"]])
                vis.info(f"  Остаток {src_len}м x {src_w}мм: вариант #{idx} -> [{pattern_desc}]")

        all_primary_cuts = OPT_CASCADING_PLAN.get("primary_cuts", [])
        solid_cuts = [cut for cut in all_primary_cuts if cut["rest"] == 0]

        for cut in solid_cuts:
            length = cut["lengths"][0] if cut.get("lengths") else 6.0
            width_mm = cut["width"]
            lc = ls.normalize_load_code(cut.get("load_code", 8))
            cut["reinforcement"] = (
                get_reinforcement_from_map(
                    reinforcement_map, length, width_mm, lc, hypothesis_debug_log=traces_p.debug_log
                )
                or 999.0
            )
        solid_cuts.sort(key=lambda x: (x.get("reinforcement", 999.0), -x["lengths"][0] if x.get("lengths") else 0))

        cut_with_rest_raw = [cut for cut in all_primary_cuts if cut["rest"] > 0]
        for cut in cut_with_rest_raw:
            length = cut["lengths"][0] if cut.get("lengths") else 6.0
            width_mm = cut["width"]
            lc = ls.normalize_load_code(cut.get("load_code", 8))
            cut["reinforcement"] = (
                get_reinforcement_from_map(
                    reinforcement_map, length, width_mm, lc, hypothesis_debug_log=traces_p.debug_log
                )
                or 999.0
            )

        cut_with_rest = sorted(
            cut_with_rest_raw,
            key=lambda x: (x.get("reinforcement", 999.0), x["width"], x["rest"]),
        )

        vis.info(f"[VISUAL] Разделение: {len(solid_cuts)} типов целых плит, {len(cut_with_rest)} типов с резом")
        if solid_cuts:
            vis.info(
                f"[VISUAL] Целые плиты (сортировка по армированию): "
                f"{[(c['width'], c['qty'], c.get('reinforcement', '?')) for c in solid_cuts[:5]]}"
            )

        logger = logging.getLogger(__name__)
        logger.info("[TRACE] ===== ШАГ 3: ПЛИТЫ ИЗ ОПТИМИЗАЦИИ (primary_cuts) =====")
        total_from_primary = sum(c["qty"] for c in all_primary_cuts)
        logger.info(f"[TRACE] Всего записей primary_cuts: {len(all_primary_cuts)}")
        logger.info(f"[TRACE] Всего плит: {total_from_primary}")

        for idx, cut in enumerate(all_primary_cuts):
            lengths_str = ", ".join([f"{lng:.2f}" for lng in cut.get("lengths", [])[:3]])
            if len(cut.get("lengths", [])) > 3:
                lengths_str += f", ... ({len(cut['lengths'])} шт)"
            logger.info(
                f"[TRACE]   #{idx+1}: width={cut['width']}мм, rest={cut['rest']}мм, qty={cut['qty']}, "
                f"lengths=[{lengths_str}], kp_id={cut.get('kp_id', '?')}"
            )

        cut_groups = [
            list(group)
            for _k, group in groupby(
                cut_with_rest, key=lambda x: (x["width"], x["rest"], x.get("reinforcement", 999.0))
            )
        ]

        if cut_groups:
            vis.info(
                f"[VISUAL] Найдено {len(cut_groups)} групп резов (сгруппировано по рез+армирование):"
            )
            for i, grp in enumerate(cut_groups, 1):
                vis.info(
                    f"[VISUAL]   Группа {i}: width={grp[0]['width']}мм, rest={grp[0]['rest']}мм, "
                    f"армирование={grp[0].get('reinforcement', '?'):.1f}, плит={sum(c['qty'] for c in grp)}"
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

        vis.info(f"[VISUAL] Развёрнуто {len(solid_cuts_list)} отдельных целых плит для разделителей")

        if solid_cuts_list:
            first_plate = solid_cuts_list.pop(0)
            ordered_cuts.append(first_plate)
            first_width = first_plate.get("width", 1200)
            vis.info(f"[VISUAL] ✓ Первая плита: целая {first_width}мм")

        for i, cut_group in enumerate(cut_groups):
            total_group_length = sum(
                cut["lengths"][0] * cut["qty"] for cut in cut_group if cut.get("lengths")
            )
            vis.info(
                f"[VISUAL] Группа резов #{i+1}: width={cut_group[0]['width']}мм, rest={cut_group[0]['rest']}мм, "
                f"типов={len(cut_group)}, длина={total_group_length:.1f}м"
            )

            if total_group_length > 90.0:
                vis.info(
                    f"[VISUAL] ⚠️ Группа #{i+1} слишком большая ({total_group_length:.1f}м > 90м), разбиваем на подгруппы"
                )
                subgroups = split_group_into_subgroups(cut_group, max_length=90.0, log=deps.log)

                for j, subgroup in enumerate(subgroups):
                    ordered_cuts.extend(subgroup)
                    vis.info(f"[VISUAL]   Добавлена подгруппа #{i+1}.{j+1}: {len(subgroup)} плит")

                    if j < len(subgroups) - 1 and solid_cuts_list:
                        separator = solid_cuts_list.pop(0)
                        separator["is_separator"] = True
                        ordered_cuts.append(separator)
                        vis.info("[VISUAL]   ✓ Разделитель между подгруппами: целая плита")
            else:
                ordered_cuts.extend(cut_group)
                vis.info(f"[VISUAL] Добавлена группа резов #{i+1} (влезает в дорожку)")

            if i < len(cut_groups) - 1 and solid_cuts_list:
                next_group = cut_groups[i + 1]
                best_idx = choose_best_separator(solid_cuts_list, next_group, reinforcement_map, log=deps.log)

                if best_idx is not None:
                    separator = solid_cuts_list.pop(best_idx)
                    separator["is_separator"] = True
                    ordered_cuts.append(separator)
                    vis.info("[VISUAL] ✓ Разделитель (is_separator=True): целая плита между группами")
                else:
                    if solid_cuts_list:
                        fallback_sep = solid_cuts_list.pop(0)
                        fallback_sep["is_separator"] = True
                        ordered_cuts.append(fallback_sep)
                        vis.info(
                            "[VISUAL] ✓ Разделитель: целая плита между группами (fallback, is_separator=True)"
                        )

        if solid_cuts_list:
            ordered_cuts.extend(solid_cuts_list)
            vis.info(f"[VISUAL] Добавлено {len(solid_cuts_list)} оставшихся целых плит в конец")

        logger.info("[TRACE] ===== ШАГ 4: ПЛИТЫ ПОСЛЕ ГРУППИРОВКИ (ordered_cuts) =====")
        total_ordered = sum(c["qty"] for c in ordered_cuts)
        logger.info(f"[TRACE] Всего записей: {len(ordered_cuts)}")
        logger.info(f"[TRACE] Всего плит: {total_ordered}")

        if total_from_primary != total_ordered:
            logger.warning("[WARNING] Потеря плит на этапе группировки!")
            logger.warning(f"[WARNING]   До группировки: {total_from_primary}")
            logger.warning(f"[WARNING]   После группировки: {total_ordered}")
            logger.warning(f"[WARNING]   Потеряно: {total_from_primary - total_ordered}")

        for idx, cut in enumerate(ordered_cuts):
            is_sep = " [РАЗДЕЛИТЕЛЬ]" if cut.get("is_separator") else ""
            rf = cut.get("reinforcement", "?")
            rf_s = f"{rf:.1f}" if isinstance(rf, (int, float)) else str(rf)
            logger.info(
                f"[TRACE]   #{idx+1}: width={cut['width']}мм, rest={cut['rest']}мм, qty={cut['qty']}, "
                f"reinf={rf_s}{is_sep}"
            )

        _missing_reinf_logged_w: set[tuple[Any, Any, Any]] = set()

        def _warn_missing_reinforcement(length: Any, width_mm: Any, load_code_from_cut: Any) -> None:
            if load_code_from_cut is None:
                return
            mw_key = (round(length, 3), width_mm, load_code_from_cut)
            if mw_key in _missing_reinf_logged_w:
                return
            _missing_reinf_logged_w.add(mw_key)
            logger.warning(
                "[ВИЗУАЛИЗАЦИЯ] Нет армирования в карте для (length=%s, width_mm=%s, load_code=%s)",
                length,
                width_mm,
                load_code_from_cut,
            )

        for cut in ordered_cuts:
            width_mm = cut["width"]
            rest_mm = cut["rest"]
            qty = cut["qty"]

            if use_2d_data and "lengths" in cut:
                lengths_for_cut = cut["lengths"]
                vis.info(f"[VISUAL] Первичный рез {width_mm}мм: используем точные длины {lengths_for_cut}")
            else:
                matching_plates = [p for p in all_plates_with_lengths if p["width"] == width_mm]
                lengths_for_cut = [p["length"] for p in matching_plates[:qty]]
                while len(lengths_for_cut) < qty:
                    lengths_for_cut.append(6.0 if not matching_plates else matching_plates[0]["length"])

            for i in range(qty):
                length = lengths_for_cut[i] if i < len(lengths_for_cut) else 6.0

                kp_id = cut.get("kp_id")
                customer = cut.get("customer")
                kp_date = cut.get("kp_date")
                plate_name_from_cut = cut.get("plate_name")

                sec_variants = secondary_cuts_info.get(secondary_geom_cut_key(length, rest_mm)) or []

                if rest_mm > 0:
                    found_variant = any(variant["used"] < variant["qty"] for variant in sec_variants)
                    vis.info(
                        f"[VISUAL] Ищем вторичные резы для остатка {length}м x {rest_mm}мм: "
                        f"{'НАЙДЕНО' if found_variant else 'НЕ НАЙДЕНО'}"
                    )

                transverse_cut_info = transverse_cut_map.get((length, width_mm))

                if transverse_cut_info:
                    width_m = width_mm / 1000.0
                    target_length = transverse_cut_info["target_length"]
                    remainder = transverse_cut_info["remainder"]

                    load_code_from_cut = ls.normalize_load_code(cut.get("load_code", 8))
                    reinforcement = get_reinforcement_from_map(
                        reinforcement_map,
                        length,
                        width_mm,
                        load_code_from_cut,
                        hypothesis_debug_log=traces_p.debug_log,
                    )
                    if reinforcement is None:
                        _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                    sequence.append(
                        {
                            "length": length,
                            "mode": "transverse",
                            "target_length": target_length,
                            "remainder": remainder,
                            "width": width_m,
                            "load_code": load_code_from_cut,
                            "label_target": plate_label(target_length, width_m, load_code_from_cut),
                            "label_remainder": f"Остаток {remainder:.2f}м".replace(".", ",")
                            if remainder > 0.1
                            else "",
                            "reinforcement": reinforcement,
                            "kp_id": kp_id,
                            "customer": customer,
                            "kp_date": kp_date,
                            "plate_name": plate_name_from_cut,
                            "concrete_grade": cut.get("concrete_grade"),
                        }
                    )
                    vis.info(
                        f"[VISUAL] Плита с поперечным резом: {length}м x {width_mm}мм -> {target_length}м "
                        f"(остаток {remainder:.2f}м)"
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
                        load_code_from_cut = ls.normalize_load_code(cut.get("load_code", 8))
                        reinforcement = get_reinforcement_from_map(
                            reinforcement_map,
                            length,
                            width_mm,
                            load_code_from_cut,
                            hypothesis_debug_log=traces_p.debug_log,
                        )
                        if reinforcement is None:
                            _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                        if abs(length - 7.1) < 0.05 and load_code_from_cut in (12, 12.5, 13):
                            try:
                                append_json_line(
                                    deps.traces.debug_log,
                                    {
                                        "hypothesisId": "H71item",
                                        "location": "layout_sequence:sequence_append_solid",
                                        "message": "71-12п item в sequence",
                                        "data": {
                                            "length": length,
                                            "width_mm": width_mm,
                                            "load_code_from_cut": load_code_from_cut,
                                            "reinforcement": reinforcement,
                                        },
                                        "timestamp": time.time() * 1000,
                                    },
                                    ensure_ascii=False,
                                )
                            except Exception:
                                pass

                        is_separator = cut.get("is_separator", False)
                        sequence.append(
                            {
                                "length": length,
                                "mode": "solid",
                                "width": main_w,
                                "load_code": load_code_from_cut,
                                "label": plate_label(length, main_w, load_code_from_cut),
                                "reinforcement": reinforcement,
                                "is_separator": is_separator,
                                "kp_id": kp_id,
                                "customer": customer,
                                "kp_date": kp_date,
                                "plate_name": plate_name_from_cut,
                                "concrete_grade": cut.get("concrete_grade"),
                            }
                        )
                    else:
                        load_code_from_cut = ls.normalize_load_code(cut.get("load_code", 8))
                        secondary_cuts_for_plate = None
                        chosen_variant = None
                        for variant in sec_variants:
                            if variant["used"] < variant["qty"]:
                                chosen_variant = variant
                                break

                        if chosen_variant:
                            secondary_cuts_for_plate = []
                            _sec_tok = chosen_variant.get("target_order_key")
                            for sec_cut_template in chosen_variant["pattern"]:
                                sec_width = sec_cut_template["width"]
                                sec_width_mm = sec_cut_template["width_mm"]
                                lc_s = sec_cut_template.get("target_load_code")
                                lc_use = ls.normalize_load_code(lc_s) if lc_s is not None else load_code_from_cut

                                sec_transverse = transverse_cut_map.get((length, sec_width_mm))

                                if sec_transverse:
                                    secondary_cuts_for_plate.append(
                                        {
                                            "width": sec_width,
                                            "label": (
                                                f'[2] {plate_label(sec_transverse["target_length"], sec_width, lc_use)}'
                                            ),
                                            "transverse_cut": True,
                                            "target_length": sec_transverse["target_length"],
                                            "remainder": sec_transverse["remainder"],
                                            "load_code": lc_use,
                                            "target_order_key": _sec_tok,
                                        }
                                    )
                                    vis.info(
                                        f"[VISUAL] Вторичный рез С поперечным: {length}м x {sec_width_mm}мм -> "
                                        f"{sec_transverse['target_length']}м"
                                    )
                                else:
                                    target_length_s = sec_cut_template.get("target_length")

                                    if target_length_s:
                                        secondary_cuts_for_plate.append(
                                            {
                                                "width": sec_width,
                                                "label": f"О {plate_label(target_length_s, sec_width, lc_use)}",
                                                "has_transverse": True,
                                                "target_length": target_length_s,
                                                "load_code": lc_use,
                                                "target_order_key": _sec_tok,
                                            }
                                        )
                                    else:
                                        result_width = sec_cut_template["width"]
                                        source_width = (
                                            sec_cut_template.get("source_width_mm", result_width * 1000) / 1000.0
                                        )
                                        label_text = plate_label(length, result_width, lc_use)
                                        if abs(result_width - source_width) > 1e-6:
                                            label_text = f"О {label_text}"
                                        secondary_cuts_for_plate.append(
                                            {
                                                "width": result_width,
                                                "label": label_text,
                                                "load_code": lc_use,
                                                "target_order_key": _sec_tok,
                                            }
                                        )
                            chosen_variant["used"] += 1

                        reinforcement = get_reinforcement_from_map(
                            reinforcement_map,
                            length,
                            width_mm,
                            load_code_from_cut,
                            hypothesis_debug_log=traces_p.debug_log,
                        )
                        if reinforcement is None:
                            _warn_missing_reinforcement(length, width_mm, load_code_from_cut)
                        sequence.append(
                            {
                                "length": length,
                                "mode": "split",
                                "main_w": main_w,
                                "rest_w": rest_w,
                                "load_code": load_code_from_cut,
                                "label_main": plate_label(length, main_w, load_code_from_cut),
                                "label_rest": (
                                    "+0,12"
                                    if fake_rest_override
                                    else (
                                        f"+{rest_w:.2f}".replace(".", ",") if not secondary_cuts_for_plate else None
                                    )
                                ),
                                "secondary_cuts": secondary_cuts_for_plate,
                                "reinforcement": reinforcement,
                                "kp_id": kp_id,
                                "customer": customer,
                                "kp_date": kp_date,
                                "plate_name": plate_name_from_cut,
                                "concrete_grade": cut.get("concrete_grade"),
                            }
                        )

        if sequence:
            logger.info("[TRACE] ===== ШАГ 5: ФИНАЛЬНАЯ ПОСЛЕДОВАТЕЛЬНОСТЬ (sequence) =====")
            logger.info(f"[TRACE] Всего плит в sequence: {len(sequence)}")
            solid_count = sum(1 for s in sequence if s.get("mode") == "solid")
            split_count = sum(1 for s in sequence if s.get("mode") == "split")
            transverse_count = sum(1 for s in sequence if s.get("mode") == "transverse")

            logger.info(f"[TRACE] Плит solid (без реза): {solid_count}")
            logger.info(f"[TRACE] Плит split (с резом): {split_count}")
            logger.info(f"[TRACE] Плит transverse (поперечный рез): {transverse_count}")

            secondary_count = 0
            for s in sequence:
                if s.get("mode") == "split" and s.get("secondary_cuts"):
                    secondary_count += len(s["secondary_cuts"])
            logger.info(f"[TRACE] Плит из вторичных резов: {secondary_count}")

            total_in_sequence = len(sequence) + secondary_count
            logger.info(f"[TRACE] ИТОГО плит: {total_in_sequence}")

            if total_from_primary != total_in_sequence:
                logger.error("[CRITICAL] ПОТЕРЯ ПЛИТ ОБНАРУЖЕНА!")
                logger.error(f"[CRITICAL]   Запрошено из оптимизации: {total_from_primary}")
                logger.error(f"[CRITICAL]   Получено в sequence:      {total_in_sequence}")
                logger.error(f"[CRITICAL]   ПОТЕРЯНО: {total_from_primary - total_in_sequence} плит(ы)")

                requested_by_width: dict[Any, Any] = {}
                for cut in all_primary_cuts:
                    ww = cut["width"]
                    requested_by_width[ww] = requested_by_width.get(ww, 0) + cut["qty"]

                result_by_width: dict[Any, Any] = {}
                for s in sequence:
                    ww = s.get(
                        "width",
                        (s.get("main_w", 1.2) * 1000 if "main_w" in s else 1200),
                    )
                    result_by_width[ww] = result_by_width.get(ww, 0) + 1
                    for sec in s.get("secondary_cuts", []):
                        sec_w = sec.get("width", 0)
                        result_by_width[sec_w] = result_by_width.get(sec_w, 0) + 1

                logger.error("[CRITICAL] Сравнение по ширинам:")
                all_widths = set(requested_by_width.keys()) | set(result_by_width.keys())
                for w in sorted(all_widths):
                    req_w = requested_by_width.get(w, 0)
                    res_w = result_by_width.get(w, 0)
                    diff = req_w - res_w
                    if diff != 0:
                        logger.error(
                            f"[CRITICAL]   Ширина {w}мм: запрошено {req_w}, получено {res_w}, ПОТЕРЯ: {diff}"
                        )
            else:
                logger.info(f"[TRACE] ✓ Проверка пройдена: все {total_from_primary} плит в sequence")

            try:
                _target_keys = [(6.0, 1200, 8), (6.0, 530, 8), (5.1, 320, 8)]
                _seq_by_key = {tuple(tk): 0 for tk in _target_keys}
                _seq_6_530_1200: list[dict[str, Any]] = []
                for s in sequence:
                    L_m = round(float(s.get("length", 0) or s.get("target_length", 0)), 2)
                    ww = s.get("width") or s.get("main_w") or 1.2
                    w_mm = round(float(ww) * 1000) if float(ww) < 20 else round(float(ww))
                    lc = s.get("load_code", 8)
                    try:
                        lc = int(lc) if lc is not None else 8
                    except (TypeError, ValueError):
                        lc = 8
                    for tk in _target_keys:
                        if abs(L_m - tk[0]) <= 0.02 and w_mm == tk[1] and lc == tk[2]:
                            _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                            break
                    if 5.98 <= L_m <= 6.02 and w_mm in (530, 1200) and len(_seq_6_530_1200) < 25:
                        _seq_6_530_1200.append(
                            {
                                "length": L_m,
                                "width_mm": w_mm,
                                "mode": s.get("mode"),
                                "label": (s.get("label") or "")[:50],
                            }
                        )
                    for sec in s.get("secondary_cuts", []):
                        sw = sec.get("width", 0)
                        sw_mm = round(float(sw) * 1000) if float(sw) < 20 else round(float(sw))
                        sl = round(float(sec.get("target_length") or L_m), 2)
                        for tk in _target_keys:
                            if abs(sl - tk[0]) <= 0.02 and sw_mm == tk[1]:
                                _seq_by_key[tuple(tk)] = _seq_by_key.get(tuple(tk), 0) + 1
                                break

                append_json_line(
                    traces_p.debug_2d5c43,
                    {
                        "sessionId": "2d5c43",
                        "hypothesisId": "H3",
                        "location": "layout_sequence:build_layout_sequence:before_return_sequence",
                        "message": "sequence vs primary totals and by key",
                        "data": {
                            "total_from_primary": total_from_primary,
                            "total_in_sequence": len(sequence),
                            "sequence_by_key": dict(_seq_by_key),
                            "sequence_6m_530_1200_sample": _seq_6_530_1200,
                        },
                        "timestamp": time.time(),
                    },
                    ensure_ascii=False,
                )
            except Exception:
                pass

            return sequence
    else:
        vis.info("[VISUAL] ВНИМАНИЕ: OPT_CASCADING_PLAN не найден или пуст, используем старый метод")

    if OPT_PLAN and OPT_PLAN.get("actions"):
        for act in OPT_PLAN["actions"]:
            src_type, W1, W2, L, qty, lc, tc = act
            W1_m = W1 / 1000.0
            W2_m = W2 / 1000.0 if W2 else 0
            for _ in range(qty):
                if src_type == "solid":
                    sequence.append({"length": L, "mode": "solid", "label": plate_label(L, W1_m)})
                elif src_type == "split":
                    rest_w = W2_m if W2_m < W1_m else (1.2 - W1_m)
                    rest_label = f"+{rest_w:.2f}".replace(".", ",")
                    sequence.append(
                        {
                            "length": L,
                            "mode": "split",
                            "main_w": W1_m,
                            "rest_w": rest_w,
                            "label_main": plate_label(L, W1_m),
                            "label_rest": rest_label,
                        }
                    )
                elif src_type == "narrow":
                    delta = abs(W2_m - W1_m) if W2_m else 0
                    rest_label = f"-{delta:.2f}".replace(".", ",") if delta > 0.001 else ""
                    sequence.append(
                        {
                            "length": L,
                            "mode": "split",
                            "main_w": W1_m,
                            "rest_w": delta,
                            "label_main": plate_label(L, W1_m),
                            "label_rest": rest_label,
                        }
                    )
        return sequence

    for L in pl.plates_1_2:
        sequence.append({"length": L, "mode": "solid", "label": plate_label(L, 1.2)})
    for L in pl.plates_1_5_to_1_2:
        sequence.append({"length": L, "mode": "solid", "label": plate_label(L, 1.2)})
    for L in pl.plates_1_0:
        sequence.append(
            {
                "length": L,
                "mode": "split",
                "main_w": 1.0,
                "rest_w": 0.2,
                "label_main": plate_label(L, 1.0),
                "label_rest": "+0,2",
            }
        )
    for L in pl.plates_1_08:
        sequence.append(
            {
                "length": L,
                "mode": "split",
                "main_w": 1.08,
                "rest_w": 0.12,
                "label_main": plate_label(L, 1.08),
                "label_rest": "+0,12",
            }
        )

    groups_map: dict[str, Any] = {
        "0_32": (pl.plates_0_32, 0.32, 0.88, "+0,88"),
        "0_46": (pl.plates_0_46, 0.46, 0.74, "+0,74"),
        "0_70": (pl.plates_0_70, 0.70, 0.50, "+0,50"),
        "0_72": (pl.plates_0_72, 0.72, 0.48, "+0,48"),
        "0_86": (pl.plates_0_86, 0.86, 0.34, "+0,34"),
    }
    if len(pl.plates_0_74):
        groups_map["0_74"] = (pl.plates_0_74, 0.74, 0.46, "+0,46")
    if len(pl.plates_0_88):
        groups_map["0_88"] = (pl.plates_0_88, 0.88, 0.32, "+0,32")
    if len(pl.plates_0_48):
        groups_map["0_48"] = (pl.plates_0_48, 0.48, 0.72, "+0,72")
    if len(pl.plates_0_50):
        groups_map["0_50"] = (pl.plates_0_50, 0.50, 0.70, "+0,70")
    if len(pl.plates_0_34):
        groups_map["0_34"] = (pl.plates_0_34, 0.34, 0.86, "+0,86")

    order_keys = OPT_WIDTH_PRIORITY or list(groups_map.keys())
    for gkey in order_keys:
        if gkey not in groups_map:
            continue
        items_t, main_w, rest_w, rest_label = groups_map[gkey]
        for L in items_t:
            sequence.append(
                {
                    "length": L,
                    "mode": "split",
                    "main_w": main_w,
                    "rest_w": rest_w,
                    "label_main": plate_label(L, main_w),
                    "label_rest": rest_label,
                }
            )

    return sequence
