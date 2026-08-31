# -*- coding: utf-8 -*-
"""Оркестрация build_layout_sequence: план, группы нагрузок, fallback-ветки."""
from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import TYPE_CHECKING, Any

from viz_modules.layout_sequence.debug_trace import LayoutSequenceTracePaths, append_json_line
from viz_modules.layout_sequence.deps import LayoutSequenceDeps
from viz_modules.layout_sequence.from_plan import _build_sequence_from_plan
from viz_modules.layout_sequence.helpers import ensure_sequence_layout_uid
from viz_modules.layout_sequence.reinforcement_data import build_reinforcement_map_for_sequence

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
