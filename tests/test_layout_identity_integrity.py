from __future__ import annotations

from app.services.production_planning_service import ProductionPlanningService
from core.visualization import validate_track_integrity
from viz_modules.layout_sequence import _build_sequence_from_plan


def test_validate_track_integrity_detects_missing_and_duplicated() -> None:
    sequence = [
        {"length": 6.0, "mode": "solid", "layout_uid": "u-1"},
        {"length": 6.0, "mode": "solid", "layout_uid": "u-2"},
    ]
    tracks = [
        {"items": [{"length": 6.0, "mode": "solid", "layout_uid": "u-1"}]},
        {"items": [{"length": 6.0, "mode": "solid", "layout_uid": "u-1"}]},
    ]

    report = validate_track_integrity(sequence, tracks, strict=False)

    assert report["ok"] is False
    assert report["missing"] == {"u-2": 1}
    assert report["duplicated"] == {"u-1": 1}


def test_build_sequence_prefers_parent_instance_id_over_legacy_key() -> None:
    plan = {
        "plate_assignments": [{"source": "primary"}],
        "primary_cuts": [
            {
                "width": 1200,
                "rest": 0,
                "qty": 1,
                "lengths": [6.0],
                "load_code": 8,
                "primary_instance_id": "prim-solid",
            },
            {
                "width": 320,
                "rest": 880,
                "qty": 1,
                "lengths": [6.0],
                "load_code": 8,
                "primary_instance_id": "prim-a",
            },
            {
                "width": 320,
                "rest": 880,
                "qty": 1,
                "lengths": [6.0],
                "load_code": 8,
                "primary_instance_id": "prim-b",
            },
        ],
        "secondary_cuts": [
            {
                "source": 999,  # намеренно не совпадает с rest=880
                "cuts": [320],
                "qty": 1,
                "pieces": 1,
                "source_lengths": [6.0],
                "lengths": [6.0],
                "target_order_key": (6.0, 320, 8),
                "parent_instance_id": "prim-b",
                "secondary_instance_id": "sec-b-1",
            }
        ],
    }

    seq = _build_sequence_from_plan(
        plan,
        plate_label_func=lambda l, w, load_code=None: f"{l}-{w}-{load_code}",
        reinforcement_map={},
    )
    split_items = [item for item in seq if item.get("mode") == "split"]
    by_unit_id = {item.get("unit_id"): item for item in split_items}

    assert by_unit_id["prim-b"]["secondary_cuts"]
    sec_item = by_unit_id["prim-b"]["secondary_cuts"][0]
    assert sec_item["unit_id"] == "sec-b-1"
    assert sec_item["parent_unit_id"] == "prim-b"
    assert by_unit_id["prim-a"].get("secondary_cuts") in (None, [])


def test_build_sequence_geometric_fallback_when_parent_instance_mismatch() -> None:
    """Вторичка с чужим parent_instance_id всё же привязывается по ключу (length, rest)."""
    plan = {
        "plate_assignments": [{"source": "primary"}],
        "primary_cuts": [
            {
                "width": 1200,
                "rest": 0,
                "qty": 1,
                "lengths": [6.0],
                "load_code": 8,
                "primary_instance_id": "prim-solid",
            },
            {
                "width": 320,
                "rest": 880,
                "qty": 1,
                "lengths": [6.0],
                "load_code": 8,
                "primary_instance_id": "prim-on-split",
            },
        ],
        "secondary_cuts": [
            {
                "source": 880,
                "cuts": [320],
                "qty": 1,
                "pieces": 1,
                "source_lengths": [6.0],
                "lengths": [6.0],
                "target_order_key": (6.0, 320, 8),
                "parent_instance_id": "wrong-parent",
                "secondary_instance_id": "sec-recovered",
            }
        ],
    }

    seq = _build_sequence_from_plan(
        plan,
        plate_label_func=lambda l, w, load_code=None: f"{l}-{w}-{load_code}",
        reinforcement_map={},
    )
    split_items = [item for item in seq if item.get("mode") == "split"]
    assert len(split_items) == 1
    split_item = split_items[0]
    assert split_item.get("unit_id") == "prim-on-split"
    assert split_item.get("secondary_cuts")
    sec_item = split_item["secondary_cuts"][0]
    assert sec_item["unit_id"] == "sec-recovered"
    assert sec_item["parent_unit_id"] == "prim-on-split"


def test_build_assignment_gap_fallback_tracks_uses_unit_identity() -> None:
    assignments = [
        {
            "unit_id": "prim-1",
            "source": "primary",
            "length": 6.0,
            "width": 1200,
            "kp_id": 1,
            "plate_name": "A",
            "load_code": 8,
        },
        {
            "unit_id": "sec-1",
            "parent_unit_id": "prim-1",
            "source": "secondary",
            "length": 6.0,
            "width": 320,
            "kp_id": 1,
            "plate_name": "A-SEC",
            "load_code": 8,
        },
    ]
    tracks = [
        {
            "items": [
                {"unit_id": "prim-1", "mode": "solid", "length": 6.0, "label": "A"}
            ]
        }
    ]

    fallback_tracks, missing = ProductionPlanningService._build_assignment_gap_fallback_tracks(
        plate_assignments=assignments,
        tracks_list=tracks,
    )

    assert len(fallback_tracks) == 1
    fallback_item = fallback_tracks[0]["items"][0]
    assert fallback_item["unit_id"] == "sec-1"
    assert fallback_item["placement_status"] == "fallback"
    assert fallback_item["mode"] == "split"
    assert sum(missing.values()) == 1
