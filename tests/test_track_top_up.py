# -*- coding: utf-8 -*-
"""Тесты дозаполнения хвоста дорожек из последующих дорожек."""

from __future__ import annotations

import copy

import pytest

from core.config.constants import TRACK_LENGTH_M
from core.track_top_up import (
    is_eligible_donor,
    recalc_track,
    top_up_tracks_from_following,
)


def _solid(
    length: float,
    *,
    reinforcement: float = 30.0,
    load_code: int = 8,
    width: float = 1.2,
    plate_uid: str | None = None,
    is_separator: bool = False,
) -> dict:
    item = {
        "length": length,
        "mode": "solid",
        "width": width,
        "load_code": load_code,
        "reinforcement": reinforcement,
        "is_separator": is_separator,
        "label": f"Плита {length}м",
    }
    if plate_uid:
        item["plate_uid"] = plate_uid
    return item


def _track(*items: dict, load_code: int = 8) -> dict:
    track = {
        "items": list(items),
        "load_code": load_code,
        "label": f"Нагрузка {load_code}п",
    }
    recalc_track(track)
    return track


def test_moves_from_following_track_to_fill_gap():
    tracks = [
        _track(_solid(90.0, plate_uid="a"), _solid(5.0, plate_uid="b")),
        _track(
            _solid(8.0, plate_uid="c0"),
            _solid(3.0, plate_uid="c1"),
            _solid(4.0, plate_uid="c2"),
        ),
    ]
    result = top_up_tracks_from_following(tracks)
    assert result.moves >= 1
    assert tracks[0]["length"] > 95.0
    assert tracks[0]["length"] <= TRACK_LENGTH_M + 0.01
    uids = [
        item.get("plate_uid")
        for track in tracks
        for item in track["items"]
        if item.get("plate_uid")
    ]
    assert len(uids) == len(set(uids))


@pytest.mark.parametrize(
    "item,index,expect",
    [
        (_solid(5.0), 0, False),
        (_solid(5.0, is_separator=True), 1, False),
        ({**_solid(5.0), "mode": "split", "secondary_cuts": [{}]}, 1, False),
        ({**_solid(5.0), "mode": "transverse"}, 1, False),
        (_solid(5.0, reinforcement=50.0), 1, False),
        (_solid(5.0, reinforcement=30.0), 1, True),
        (_solid(120.0, reinforcement=30.0), 1, False),
    ],
)
def test_is_eligible_donor_rules(item, index, expect):
    assert (
        is_eligible_donor(item, index, gap=10.0, max_reinforcement=40.0) is expect
    )


def test_does_not_take_from_same_or_previous_track():
    tracks = [
        _track(_solid(90.0, plate_uid="t0a"), _solid(5.0, plate_uid="t0b")),
        _track(_solid(8.0, plate_uid="t1a"), _solid(2.0, plate_uid="t1b")),
    ]
    before = copy.deepcopy(tracks)
    top_up_tracks_from_following(tracks)
    assert tracks[0]["items"][0]["plate_uid"] == "t0a"
    assert tracks[1]["items"][0]["plate_uid"] == "t1a"


def test_cross_load_code_allowed_when_reinforcement_ok():
    tracks = [
        _track(_solid(95.0, reinforcement=40.0, load_code=10, plate_uid="recv")),
        _track(
            _solid(8.0, reinforcement=35.0, load_code=8, plate_uid="donor0"),
            _solid(5.0, reinforcement=38.0, load_code=8, plate_uid="donor1"),
        ),
    ]
    result = top_up_tracks_from_following(tracks)
    assert result.moves == 1
    assert tracks[0]["items"][-1]["plate_uid"] == "donor1"
    assert tracks[0]["length"] == pytest.approx(100.0, abs=0.01)


def test_prefers_longer_higher_reinforcement_candidate():
    tracks = [
        _track(_solid(96.0, reinforcement=40.0, plate_uid="recv")),
        _track(
            _solid(8.0, plate_uid="d0"),
            _solid(3.0, reinforcement=35.0, plate_uid="short"),
            _solid(5.0, reinforcement=39.0, plate_uid="long"),
        ),
    ]
    top_up_tracks_from_following(tracks)
    moved_uids = [item.get("plate_uid") for item in tracks[0]["items"]]
    assert "long" in moved_uids
    assert "short" not in moved_uids


def test_respects_max_length_limit():
    tracks = [
        _track(_solid(99.5, reinforcement=40.0, plate_uid="recv")),
        _track(
            _solid(8.0, plate_uid="d0"),
            _solid(2.0, reinforcement=30.0, plate_uid="too_long"),
        ),
    ]
    result = top_up_tracks_from_following(tracks)
    assert result.moves == 0
    assert tracks[0]["length"] == pytest.approx(99.5, abs=0.01)


def test_donor_first_plate_remains_solid():
    tracks = [
        _track(_solid(95.0, reinforcement=40.0, plate_uid="recv")),
        _track(
            _solid(8.0, plate_uid="first"),
            _solid(4.0, reinforcement=35.0, plate_uid="tail"),
        ),
    ]
    top_up_tracks_from_following(tracks)
    assert tracks[1]["items"][0]["mode"] == "solid"
    assert tracks[1]["items"][0]["plate_uid"] == "first"
