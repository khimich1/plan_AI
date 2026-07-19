# -*- coding: utf-8 -*-
"""Инвариант: дорожка начинается с целой; плиты не теряются (strict integrity)."""

from __future__ import annotations

import pytest

from core.visualization import TrackLayoutInvariantError, split_sequence_into_tracks


def _split_item(length_m: float) -> dict:
    return {
        "length": length_m,
        "mode": "split",
        "main_w": 0.72,
        "rest_w": 0.48,
        "label_main": "S",
        "load_code": 8,
        "reinforcement": 5.0,
    }


def _solid(length_m: float, reinforcement: float = 1.0) -> dict:
    return {
        "length": length_m,
        "mode": "solid",
        "width": 1.2,
        "load_code": 8,
        "label": f"{length_m:g}m solid",
        "reinforcement": reinforcement,
    }


def test_splits_only_raises_track_invariant_flat() -> None:
    seq = [_split_item(5.0)]
    with pytest.raises(TrackLayoutInvariantError):
        split_sequence_into_tracks(
            seq,
            strict_layout_integrity=True,
            track_start_reinf_relaxation=False,
        )


def test_partner_length_impossible_raises() -> None:
    """Нет целой с суммой len+split <= 101."""
    grp = {
        "load_code": 8,
        "label": "grp",
        "sequence": [
            _solid(98.0, 1.0),
            _split_item(99.0),
            _solid(98.0, 1.0),
        ],
    }
    with pytest.raises(TrackLayoutInvariantError):
        split_sequence_into_tracks([grp], strict_layout_integrity=True)


def test_reinf_relaxation_allows_heavier_starter() -> None:
    """Фаза 2: только «тяжёлая» целая доступна после строго отсечения по арм."""
    grp = {
        "load_code": 8,
        "label": "grp",
        "sequence": [
            _solid(50.0, 100.0),
            _solid(48.0, 100.0),
            _split_item(52.0),
            _solid(2.0, 150.0),
        ],
    }
    tracks = split_sequence_into_tracks(
        [grp],
        strict_layout_integrity=True,
        track_start_reinf_relaxation=True,
    )
    assert tracks[1]["items"][0]["mode"] == "solid"
    assert sum(len(t["items"]) for t in tracks) == 4


def test_reinf_relaxation_off_raises_when_strict_blocks() -> None:
    grp = {
        "load_code": 8,
        "label": "grp",
        "sequence": [
            _solid(50.0, 100.0),
            _solid(48.0, 100.0),
            _split_item(52.0),
            _solid(2.0, 150.0),
        ],
    }
    with pytest.raises(TrackLayoutInvariantError):
        split_sequence_into_tracks(
            [grp],
            strict_layout_integrity=True,
            track_start_reinf_relaxation=False,
        )
