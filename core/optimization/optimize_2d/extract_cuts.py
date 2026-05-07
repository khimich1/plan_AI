# -*- coding: utf-8 -*-
"""Phase B: extract PuLP solution into planned primary/secondary cuts (ordering + parents)."""

from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass

from pulp import value

from core import config_and_data as cfg
from core.optimization.geometry import _canonical_length
from core.optimization.ilp_model import _residual_phys_key
from core.optimization.optimization_debug_impl import _DEBUG_LOG_7E420E
from core.optimization.optimize_2d.state import TwoDPhaseAState
from core.optimization.secondary_batches import _batch_sizes_for_secondary_z_sec


@dataclass
class TwoDPhaseBResult:
    """Outputs of solver extraction before post-correction / attribution."""

    primary_cuts: list
    secondary_cuts: list
    total_plates: int
    rests_used: list
    next_primary_instance_id: int
    next_secondary_instance_id: int
    n_solid_primary_plates: int
    n_cut_primary_plates: int


def extract_two_d_phase_b(phase_state: TwoDPhaseAState) -> TwoDPhaseBResult:
    """
    Build primary_cuts / secondary_cuts from solved ILP variables.

    Preserves primary row order after factory grouping (solids first, then cuts).
    Secondary rows use `_batch_sizes_for_secondary_z_sec` so parent_instance_id matches
    prior batched-parent semantics.
    """
    ilp = phase_state.ilp
    z_prim = ilp.z_prim
    z_sec = ilp.z_sec
    x_sec = ilp.x_sec
    primary_options_by_id = ilp.primary_options_by_id
    secondary_options_by_id = ilp.secondary_options_by_id
    secondary_options = phase_state.secondary_options

    _next_primary_instance_id = 1
    _next_secondary_instance_id = 1
    _primary_instances_by_opt_id: dict[int, list[str]] = defaultdict(list)
    _primary_instances_by_geom_lc: dict[
        tuple[float, int], list[tuple[float | int, int, str]]
    ] = defaultdict(list)

    planned_primary_cuts: list = []
    planned_secondary_cuts: list = []
    total_plates = 0

    for (opt_id, dk), zv in z_prim.items():
        raw_val = value(zv) or 0
        qty = math.ceil(raw_val - 1e-6) if raw_val > 1e-6 else 0
        if qty <= 0:
            continue
        opt = primary_options_by_id[opt_id]
        target_length, target_width, target_load_code = dk
        for _ in range(qty):
            primary_instance_id = f"prim-{_next_primary_instance_id}"
            _next_primary_instance_id += 1
            planned_primary_cuts.append(
                {
                    "width": opt["main"],
                    "demand_width": target_width,
                    "rest": opt["rest"],
                    "qty": 1,
                    "lengths": [opt["length"]],
                    "load_code": target_load_code,
                    "assignment_key": dk,
                    "source_opt_id": opt_id,
                    "primary_instance_id": primary_instance_id,
                }
            )
            _primary_instances_by_opt_id[opt_id].append(primary_instance_id)
            if opt.get("rest", 0) > 0:
                _primary_instances_by_geom_lc[
                    _residual_phys_key(opt["length"], opt["rest"])
                ].append(
                    (
                        cfg.normalize_load_code(
                            opt.get("load_code", target_load_code), default=8
                        ),
                        opt_id,
                        primary_instance_id,
                    )
                )
            total_plates += 1

    # #region agent log
    try:
        import json as _aj
        import time as _at

        _geom_prim_counts: dict[str, int] = {}
        for _pc in planned_primary_cuts:
            if _pc.get("rest", 0) <= 0:
                continue
            _L0 = (
                _canonical_length(_pc["lengths"][0]) if _pc.get("lengths") else 0.0
            )
            _rk = f"{_L0}_{int(round(float(_pc['rest'])))}"
            _geom_prim_counts[_rk] = _geom_prim_counts.get(_rk, 0) + 1
        _opt_queue_lens_before_sec = {
            str(k): len(v)
            for k, v in _primary_instances_by_opt_id.items()
            if v
        }
        with open(_DEBUG_LOG_7E420E, "a", encoding="utf-8") as _lf:
            _lf.write(
                _aj.dumps(
                    {
                        "sessionId": "7e420e",
                        "hypothesisId": "H_OPT_PRIMARY_GEOM",
                        "location": "core/optimization/optimize_2d/extract_cuts.py",
                        "message": "primary splits count by (len_m, rest_mm); opt_id queues before any secondary pop",
                        "data": {
                            "n_planned_primary": len(planned_primary_cuts),
                            "geom_split_counts": _geom_prim_counts,
                            "opt_id_queue_nonempty": _opt_queue_lens_before_sec,
                        },
                        "timestamp": int(_at.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass
    # #endregion

    print("[OPT_2D] 🔧 Применяем правила завода для порядка плит...")

    solid_plates: list = []
    cut_plates: list = []

    for cut in planned_primary_cuts:
        if cut["rest"] == 0:
            solid_plates.append(cut)
        else:
            cut_plates.append(cut)

    solid_plates.sort(
        key=lambda x: (-x["width"], -x["lengths"][0] if x.get("lengths") else 0)
    )
    cut_plates.sort(key=lambda x: (-x["rest"], -x["width"]))

    primary_cuts = solid_plates + cut_plates

    rests_used: list = []
    for opt in secondary_options:
        apps = int(round(value(x_sec[opt["id"]]) or 0))
        for _ in range(apps):
            rests_used.append(
                {
                    "source_length": opt["source_length"],
                    "source_rest_mm": opt["source_rest"],
                }
            )

    def _remove_primary_instance_from_geom(instance_id: str) -> None:
        for pool in _primary_instances_by_geom_lc.values():
            for idx, (_prim_lc, _opt_id, _inst_id) in enumerate(pool):
                if _inst_id == instance_id:
                    del pool[idx]
                    return

    _orphan_recovered_geometry = 0
    _secondary_parent_missing = 0

    for (opt_id, dk), zv in z_sec.items():
        raw_val = value(zv) or 0
        qty = int(round(raw_val))
        if qty <= 0:
            continue
        opt = secondary_options_by_id[opt_id]
        _target_length, _target_width, target_load_code = dk
        target_load_code = cfg.normalize_load_code(target_load_code, default=8)
        pieces = max(1, int(opt.get("pieces") or 1))
        if qty % pieces != 0:
            import logging as _batch_qty_log

            _batch_qty_log.getLogger(__name__).warning(
                "[OPT_2D] z_sec qty=%d not divisible by pieces=%d for sec_opt_id=%s dk=%s — "
                "last chunk batched as one parental rest",
                qty,
                pieces,
                opt_id,
                list(dk) if isinstance(dk, (list, tuple)) else dk,
            )

        def _pop_parent_for_secondary_rest() -> str | None:
            nonlocal _orphan_recovered_geometry
            parent_id: str | None = None
            for source_opt_id in opt.get("source_ids") or []:
                queue = _primary_instances_by_opt_id.get(source_opt_id) or []
                if queue:
                    parent_id = queue.pop(0)
                    _remove_primary_instance_from_geom(parent_id)
                    break
            if not parent_id:
                pool = (
                    _primary_instances_by_geom_lc.get(
                        _residual_phys_key(
                            opt.get("source_length"), opt.get("source_rest")
                        )
                    )
                    or []
                )
                for idx, (prim_lc, source_opt_id, instance_id) in enumerate(pool):
                    if cfg.normalize_load_code(prim_lc, default=8) >= target_load_code:
                        parent_id = instance_id
                        del pool[idx]
                        opt_queue = _primary_instances_by_opt_id.get(source_opt_id) or []
                        if instance_id in opt_queue:
                            opt_queue.remove(instance_id)
                        _orphan_recovered_geometry += 1
                        break
            return parent_id

        batch_sizes_list = _batch_sizes_for_secondary_z_sec(qty, pieces)
        z_block_offset = 0
        for batch_index, batch_size in enumerate(batch_sizes_list):
            _q_before = {
                str(_soid): len(_primary_instances_by_opt_id.get(_soid) or [])
                for _soid in (opt.get("source_ids") or [])
            }
            parent_instance_id = _pop_parent_for_secondary_rest()
            _q_mid = {
                str(_soid): len(_primary_instances_by_opt_id.get(_soid) or [])
                for _soid in (opt.get("source_ids") or [])
            }
            if not parent_instance_id:
                _secondary_parent_missing += batch_size
            for _ in range(batch_size):
                secondary_instance_id = f"sec-{_next_secondary_instance_id}"
                _next_secondary_instance_id += 1
                # #region agent log
                try:
                    import json as _aj
                    import time as _at

                    _q_after = {
                        str(_soid): len(_primary_instances_by_opt_id.get(_soid) or [])
                        for _soid in (opt.get("source_ids") or [])
                    }
                    with open(_DEBUG_LOG_7E420E, "a", encoding="utf-8") as _lf:
                        _lf.write(
                            _aj.dumps(
                                {
                                    "sessionId": "7e420e",
                                    "hypothesisId": "H_OPT_SEC_PARENT_POP",
                                    "location": "core/optimization/optimize_2d/extract_cuts.py",
                                    "message": "z_sec output row shares parent within pieces-batch",
                                    "data": {
                                        "sec_opt_id": opt_id,
                                        "pieces": pieces,
                                        "batch_index": batch_index,
                                        "batch_size": batch_size,
                                        "batch_offset_in_z_block": z_block_offset,
                                        "source_length": opt.get("source_length"),
                                        "source_rest": opt.get("source_rest"),
                                        "target_order_key": list(dk)
                                        if isinstance(dk, (list, tuple))
                                        else dk,
                                        "source_ids": list(opt.get("source_ids") or []),
                                        "queue_lens_before_pop": _q_before,
                                        "queue_remaining_after_parent_pop": _q_mid,
                                        "queue_remaining_by_source_opt_id": _q_after,
                                        "parent_instance_id": parent_instance_id,
                                        "secondary_instance_id": secondary_instance_id,
                                        "sec_type": opt.get("type"),
                                    },
                                    "timestamp": int(_at.time() * 1000),
                                },
                                ensure_ascii=False,
                                default=str,
                            )
                            + "\n"
                        )
                except Exception:
                    pass
                # #endregion
                planned_secondary_cuts.append(
                    {
                        "source": opt["source_rest"],
                        "cuts": [opt["output_width"]],
                        "qty": 1,
                        "pieces": 1,
                        "waste": opt.get("waste", 0),
                        "type": opt["type"],
                        "source_lengths": [opt["source_length"]],
                        "lengths": [opt["output_length"]],
                        "target_order_key": dk,
                        "load_code": target_load_code,
                        "parent_instance_id": parent_instance_id,
                        "secondary_instance_id": secondary_instance_id,
                        "source_opt_ids": list(opt.get("source_ids") or []),
                    }
                )
            z_block_offset += batch_size

    if _orphan_recovered_geometry or _secondary_parent_missing:
        import logging as _parent_log

        _parent_log.getLogger(__name__).warning(
            "[OPT_2D] secondary parent assignment: recovered_by_geometry=%d, missing=%d",
            _orphan_recovered_geometry,
            _secondary_parent_missing,
        )

    # #region agent log
    try:
        import json as _aj
        import time as _at

        _null_parent = sum(
            1 for c in planned_secondary_cuts if not c.get("parent_instance_id")
        )
        _by_geom = {}
        for c in planned_secondary_cuts:
            if c.get("parent_instance_id"):
                continue
            sl = c.get("source_lengths") or []
            _L = _canonical_length(sl[0]) if sl else None
            _src = c.get("source")
            _gk = (
                f"{_L}_{int(round(float(_src)))}"
                if _L is not None and _src is not None
                else "?"
            )
            _by_geom[_gk] = _by_geom.get(_gk, 0) + 1
        with open(_DEBUG_LOG_7E420E, "a", encoding="utf-8") as _lf:
            _lf.write(
                _aj.dumps(
                    {
                        "sessionId": "7e420e",
                        "hypothesisId": "H_OPT_SEC_SUMMARY",
                        "location": "core/optimization/optimize_2d/extract_cuts.py",
                        "message": "secondary cuts: null parent count and breakdown by geom key",
                        "data": {
                            "n_secondary": len(planned_secondary_cuts),
                            "null_parent_count": _null_parent,
                            "null_parent_by_source_geom_key": _by_geom,
                        },
                        "timestamp": int(_at.time() * 1000),
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    except Exception:
        pass
    # #endregion

    return TwoDPhaseBResult(
        primary_cuts=primary_cuts,
        secondary_cuts=planned_secondary_cuts,
        total_plates=total_plates,
        rests_used=rests_used,
        next_primary_instance_id=_next_primary_instance_id,
        next_secondary_instance_id=_next_secondary_instance_id,
        n_solid_primary_plates=len(solid_plates),
        n_cut_primary_plates=len(cut_plates),
    )
