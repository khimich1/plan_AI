# -*- coding: utf-8 -*-
"""
Три вторичные строки плана с одним parent_instance_id → один merge-variant,
все nested secondary_cuts на слоте split родителя (без dangling unit_id).
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.optimization  # noqa: E402
import viz_modules.layout_sequence as layout_sequence_mod  # noqa: E402
from viz_modules.layout_sequence import (  # noqa: E402
    _build_sequence_from_plan,
    _merge_atomic_secondaries_by_shared_parent,
)


def _minimal_label(length, width_m, lc) -> str:
    return f"{length}_{width_m}_{lc}"


def _sec_row(sec_id: str, length_m: float = 5.94) -> dict:
    lw = round(float(length_m), 2)
    return {
        "source": 900,
        "cuts": [300],
        "qty": 1,
        "pieces": 1,
        "source_lengths": [length_m],
        "lengths": [length_m],
        "target_order_key": (lw, 300, 8),
        "parent_instance_id": "prim-35",
        "secondary_instance_id": sec_id,
        "type": "multiple",
        "waste": 0,
    }


def _minimal_plan_three_secondaries_same_parent() -> dict:
    """
    Одна первичка с резом (остаток 900 мм под три полосы 300) и три строки вторичных
    с общим родителем — как после оптимизатора с батчем parent.
    """
    lm = 5.94
    return {
        "plate_assignments": [{"unit_id": "prim-35", "dummy": True}],
        "primary_cuts": [
            {
                "width": 1200,
                "rest": 900,
                "qty": 1,
                "lengths": [lm],
                "load_code": 8,
                "primary_instance_id": "prim-35",
            }
        ],
        "secondary_cuts": [
            _sec_row("sec-2", lm),
            _sec_row("sec-3", lm),
            _sec_row("sec-4", lm),
        ],
        "transverse_cuts": [],
    }


def test_merge_three_atomics_under_one_primary_all_nested_unit_ids_distinct():
    plan = _minimal_plan_three_secondaries_same_parent()
    seq = _build_sequence_from_plan(plan, _minimal_label, reinforcement_map={})
    splits = [s for s in seq if s.get("mode") == "split" and s.get("unit_id") == "prim-35"]
    assert len(splits) == 1
    nested = splits[0].get("secondary_cuts") or []
    assert len(nested) == 3
    uids = [x.get("unit_id") for x in nested]
    assert set(uids) == {"sec-2", "sec-3", "sec-4"}


def test_h5_secondary_unmatched_zero_for_merged_three():
    captured: list[dict] = []
    orig = layout_sequence_mod._agent_seq_debug

    def _spy(hypothesis_id: str, message: str, data: dict) -> None:
        if hypothesis_id == "H5" and message == "phase_end_summary":
            captured.append(dict(data))
        return orig(hypothesis_id, message, data)

    plan = _minimal_plan_three_secondaries_same_parent()
    with patch.object(layout_sequence_mod, "_agent_seq_debug", side_effect=_spy):
        _build_sequence_from_plan(plan, _minimal_label, reinforcement_map={})

    assert captured
    assert captured[-1].get("secondary_total_from_plan") == 3
    assert captured[-1].get("secondary_attached_total") == 3
    assert captured[-1].get("secondary_unmatched_total") == 0


def _seg(width_mm: int) -> dict:
    return {
        "width": width_mm / 1000.0,
        "width_mm": float(width_mm),
        "source_width_mm": 900,
        "label": None,
        "target_length": 5.94,
        "target_load_code": 8,
    }


def test_merge_helper_skips_when_segment_widths_differ():
    """Один merge_key, но разная width_mm у атомов → не склеиваем, список раздваивается."""
    import logging

    gk = (5.94, 900)
    tok = (5.94, 300, 8)
    bucket = {
        "prim-x": [
            {
                "pattern": [_seg(300)],
                "qty": 1,
                "used": 0,
                "target_order_key": tok,
                "geom_key": gk,
                "secondary_instance_id": "sec-a",
                "parent_instance_id": "prim-x",
            },
            {
                "pattern": [_seg(340)],
                "qty": 1,
                "used": 0,
                "target_order_key": tok,
                "geom_key": gk,
                "secondary_instance_id": "sec-b",
                "parent_instance_id": "prim-x",
            },
        ]
    }
    _merge_atomic_secondaries_by_shared_parent(
        secondary_cuts_by_parent=bucket,
        logger=logging.getLogger("test_merge"),
    )
    out = bucket["prim-x"]
    assert len(out) == 2
    assert all("secondary_instance_ids" in v for v in out)
    assert out[0]["secondary_instance_ids"] == ["sec-a"]
    assert out[1]["secondary_instance_ids"] == ["sec-b"]
