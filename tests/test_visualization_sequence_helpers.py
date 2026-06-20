"""Unit tests for visualization sequence flattening and grouped integrity checks."""

from __future__ import annotations

import pytest

from core.visualization import (
    LayoutIntegrityError,
    _count_solids_remaining,
    _iter_sequence_items,
    validate_track_integrity,
)


def test_iter_sequence_items_flat_passthrough() -> None:
    items = [{"mode": "solid", "length": 5.0}, {"mode": "split", "length": 3.0}]
    assert _iter_sequence_items(items) == items


def test_iter_sequence_items_flattens_grouped_sequence() -> None:
    grouped = [
        {
            "load_code": 8,
            "sequence": [
                {"mode": "solid", "length": 6.0, "layout_uid": "a"},
                {"mode": "split", "length": 4.0, "layout_uid": "b"},
            ],
        },
        {
            "load_code": 10,
            "sequence": [{"mode": "solid", "length": 5.0, "layout_uid": "c"}],
        },
    ]
    flat = _iter_sequence_items(grouped)
    assert [it["layout_uid"] for it in flat] == ["a", "b", "c"]


def test_validate_track_integrity_grouped_sequence_ok() -> None:
    sequence = [
        {
            "load_code": 8,
            "sequence": [
                {"mode": "solid", "length": 6.0, "layout_uid": "u-1"},
                {"mode": "split", "length": 4.0, "layout_uid": "u-2"},
            ],
        }
    ]
    tracks = [
        {"items": [{"mode": "solid", "length": 6.0, "layout_uid": "u-1"}]},
        {"items": [{"mode": "split", "length": 4.0, "layout_uid": "u-2"}]},
    ]

    report = validate_track_integrity(sequence, tracks, strict=False)

    assert report["ok"] is True
    assert report["missing"] == {}
    assert report["duplicated"] == {}


def test_validate_track_integrity_strict_raises_on_grouped_loss() -> None:
    sequence = [
        {
            "load_code": 8,
            "sequence": [
                {"mode": "solid", "layout_uid": "keep"},
                {"mode": "solid", "layout_uid": "lost"},
            ],
        }
    ]
    tracks = [{"items": [{"mode": "solid", "layout_uid": "keep"}]}]

    with pytest.raises(LayoutIntegrityError):
        validate_track_integrity(sequence, tracks, strict=True)


def test_count_solids_remaining_ignores_splits_and_separators() -> None:
    items = [
        {"mode": "solid"},
        {"mode": "solid", "is_separator": True},
        {"mode": "split"},
        {"mode": "transverse"},
    ]
    assert _count_solids_remaining(items) == 2
