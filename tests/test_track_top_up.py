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
    concrete_grade: str | None = None,
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
    if concrete_grade:
        item["concrete_grade"] = concrete_grade
    return item


def _bundle(tracks: list[dict]) -> tuple[list[str], float]:
    uids = [
        item["plate_uid"]
        for track in tracks
        for item in track["items"]
        if item.get("plate_uid")
    ]
    total = sum(float(track["length"]) for track in tracks)
    return uids, total


def _assert_conserved(before_uids: list[str], before_len: float, tracks: list[dict]) -> None:
    uids, total = _bundle(tracks)
    assert sorted(uids) == sorted(before_uids)
    assert len(uids) == len(set(uids))
    assert total == pytest.approx(before_len, abs=0.01)
    assert tracks
    assert all(track["items"] for track in tracks)
    assert all(track["length"] <= TRACK_LENGTH_M + 0.01 for track in tracks)


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
    "item,index,sole,expect",
    [
        (_solid(5.0), 0, False, False),
        (_solid(5.0, is_separator=True), 1, False, False),
        ({**_solid(5.0), "mode": "split", "secondary_cuts": [{}]}, 1, False, False),
        ({**_solid(5.0), "mode": "split"}, 1, False, False),
        ({**_solid(5.0), "mode": "transverse"}, 1, False, False),
        (_solid(5.0, reinforcement=50.0), 1, False, False),
        (_solid(5.0, reinforcement=30.0), 1, False, True),
        (_solid(120.0, reinforcement=30.0), 1, False, False),
        (_solid(5.0, reinforcement=30.0), 0, True, True),
    ],
)
def test_is_eligible_donor_rules(item, index, sole, expect):
    assert (
        is_eligible_donor(
            item,
            index,
            gap=10.0,
            max_reinforcement=40.0,
            receiver_grade="М500",
            sole_on_last_track=sole,
        )
        is expect
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


def test_sole_last_plate_moves_into_gap_and_empty_track_is_removed():
    tracks = [
        _track(_solid(87.7, reinforcement=30.0, concrete_grade="М400", plate_uid="body")),
        _track(_solid(2.28, reinforcement=30.0, concrete_grade="М400", plate_uid="tail")),
    ]
    before_uids, before_len = _bundle(tracks)

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 1
    assert len(tracks) == 1
    assert tracks[0]["length"] == pytest.approx(89.98, abs=0.01)
    assert [item["plate_uid"] for item in tracks[0]["items"]] == ["body", "tail"]
    _assert_conserved(before_uids, before_len, tracks)


def test_gap_4_5_takes_4m_and_leaves_2_2_track():
    tracks = [
        _track(_solid(96.5, reinforcement=40.0, plate_uid="recv")),
        _track(
            _solid(2.2, reinforcement=30.0, plate_uid="left"),
            _solid(4.0, reinforcement=30.0, plate_uid="moved"),
        ),
    ]
    before_uids, before_len = _bundle(tracks)

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 1
    assert tracks[0]["items"][-1]["plate_uid"] == "moved"
    assert tracks[0]["length"] == pytest.approx(100.5, abs=0.01)
    assert len(tracks) == 2
    assert tracks[1]["length"] == pytest.approx(2.2, abs=0.01)
    assert [item["plate_uid"] for item in tracks[1]["items"]] == ["left"]
    _assert_conserved(before_uids, before_len, tracks)


def test_shorter_than_gap_moves_and_remainder_stays_last():
    tracks = [
        _track(_solid(99.5, reinforcement=40.0, plate_uid="recv")),
        _track(
            _solid(10.0, reinforcement=30.0, plate_uid="rest"),
            _solid(1.2, reinforcement=30.0, plate_uid="short"),
        ),
    ]
    before_uids, before_len = _bundle(tracks)

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 1
    assert tracks[0]["items"][-1]["plate_uid"] == "short"
    assert tracks[0]["length"] == pytest.approx(100.7, abs=0.01)
    assert tracks[-1]["items"][0]["plate_uid"] == "rest"
    assert tracks[-1]["length"] == pytest.approx(10.0, abs=0.01)
    _assert_conserved(before_uids, before_len, tracks)


def test_index_zero_on_last_track_is_not_taken_while_another_plate_remains():
    tracks = [
        _track(_solid(96.0, reinforcement=40.0, plate_uid="recv")),
        _track(
            _solid(4.0, reinforcement=30.0, plate_uid="first"),
            _solid(3.0, reinforcement=30.0, plate_uid="second"),
        ),
    ]

    top_up_tracks_from_following(tracks)

    assert tracks[0]["items"][-1]["plate_uid"] == "second"
    assert [item["plate_uid"] for item in tracks[1]["items"]] == ["first"]


def test_sole_plate_of_middle_track_is_not_donated():
    tracks = [
        _track(_solid(90.0, reinforcement=40.0, plate_uid="recv")),
        _track(_solid(5.0, reinforcement=30.0, plate_uid="middle")),
        _track(
            _solid(20.0, plate_uid="lead"),
            _solid(3.0, reinforcement=30.0, plate_uid="tail"),
        ),
    ]

    top_up_tracks_from_following(tracks)

    assert [item["plate_uid"] for item in tracks[1]["items"]] == ["middle"]
    assert tracks[0]["items"][-1]["plate_uid"] == "tail"
    assert tracks[2]["items"][0]["plate_uid"] == "lead"


def test_higher_reinforcement_is_not_donated():
    tracks = [
        _track(_solid(90.0, reinforcement=30.0, plate_uid="recv")),
        _track(
            _solid(8.0, plate_uid="lead"),
            _solid(4.0, reinforcement=40.0, plate_uid="thick"),
        ),
    ]

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 0
    assert tracks[1]["items"][-1]["plate_uid"] == "thick"


def test_m500_by_length_does_not_move_onto_m400():
    tracks = [
        _track(_solid(6.0, reinforcement=40.0, load_code=8, plate_uid="recv")),
        _track(
            _solid(3.0, plate_uid="lead"),
            _solid(7.0, reinforcement=30.0, load_code=8, plate_uid="long"),
        ),
    ]

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 0
    assert tracks[1]["items"][-1]["plate_uid"] == "long"


def test_m400_moves_onto_track_with_m500():
    tracks = [
        _track(_solid(90.0, reinforcement=40.0, load_code=8, plate_uid="recv")),
        _track(
            _solid(8.0, plate_uid="lead"),
            _solid(4.0, reinforcement=30.0, load_code=8, plate_uid="low"),
        ),
    ]

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 1
    assert tracks[0]["items"][-1]["plate_uid"] == "low"


def test_explicit_receiver_grade_beats_length_rule():
    tracks = [
        _track(
            _solid(
                6.0,
                reinforcement=40.0,
                load_code=8,
                concrete_grade="М500",
                plate_uid="recv",
            )
        ),
        _track(
            _solid(3.0, reinforcement=50.0, plate_uid="lead"),
            _solid(7.0, reinforcement=30.0, load_code=8, plate_uid="donor"),
        ),
    ]

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 1
    assert tracks[0]["items"][-1]["plate_uid"] == "donor"


@pytest.mark.parametrize(
    "item,receiver_grade,expect",
    [
        (_solid(6.6, load_code=8, reinforcement=30.0), "М400", True),
        (_solid(6.7, load_code=8, reinforcement=30.0), "М400", False),
        (_solid(6.6, load_code=6, reinforcement=30.0), "М400", True),
        (_solid(6.7, load_code=6, reinforcement=30.0), "М400", False),
        (_solid(4.0, load_code=10, reinforcement=30.0), "М400", False),
        (_solid(4.0, load_code=8, reinforcement=30.0, concrete_grade="М500"), "М400", False),
        (_solid(8.0, load_code=8, reinforcement=30.0, concrete_grade="М400"), "М400", True),
        (_solid(4.0, load_code=8, reinforcement=30.0), "М500", True),
    ],
)
def test_donor_grade_not_above_receiver(item, receiver_grade, expect):
    assert (
        is_eligible_donor(
            item,
            1,
            gap=20.0,
            max_reinforcement=40.0,
            receiver_grade=receiver_grade,
        )
        is expect
    )


def _pair(root: dict, *cuts: dict) -> dict:
    paired = dict(root)
    paired["mode"] = "split"
    paired["secondary_cuts"] = list(cuts)
    return paired


def test_secondary_cuts_move_with_the_root_item():
    pair = _pair(
        _solid(4.0, reinforcement=30.0, load_code=8, plate_uid="pair"),
        _solid(2.0, reinforcement=28.0, load_code=8, plate_uid="half"),
    )
    tracks = [
        _track(_solid(90.0, reinforcement=40.0, plate_uid="recv")),
        _track(_solid(8.0, plate_uid="lead"), pair),
    ]
    before_uids, before_len = _bundle(tracks)

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 1
    moved = tracks[0]["items"][-1]
    assert moved["plate_uid"] == "pair"
    assert moved["secondary_cuts"][0]["plate_uid"] == "half"
    assert tracks[0]["length"] == pytest.approx(94.0, abs=0.01)
    flat = [item.get("plate_uid") for track in tracks for item in track["items"]]
    assert "half" not in flat
    _assert_conserved(before_uids, before_len, tracks)


def test_secondary_cut_reinforcement_above_receiver_blocks_item():
    pair = _pair(
        _solid(4.0, reinforcement=30.0, plate_uid="pair"),
        _solid(2.0, reinforcement=50.0, plate_uid="half"),
    )
    tracks = [
        _track(_solid(90.0, reinforcement=40.0, plate_uid="recv")),
        _track(_solid(8.0, plate_uid="lead"), pair),
    ]

    assert top_up_tracks_from_following(tracks).moves == 0
    assert tracks[1]["items"][-1]["plate_uid"] == "pair"


def test_secondary_cut_grade_above_receiver_blocks_item():
    pair = _pair(
        _solid(4.0, reinforcement=30.0, load_code=8, concrete_grade="М400", plate_uid="pair"),
        _solid(7.0, reinforcement=30.0, load_code=8, plate_uid="half"),
    )
    tracks = [
        _track(_solid(6.0, reinforcement=40.0, load_code=8, plate_uid="recv")),
        _track(_solid(3.0, plate_uid="lead"), pair),
    ]

    assert top_up_tracks_from_following(tracks).moves == 0
    assert tracks[1]["items"][-1]["secondary_cuts"][0]["plate_uid"] == "half"


def test_transverse_is_not_a_donor():
    transverse = _solid(4.0, reinforcement=30.0, plate_uid="cut")
    transverse["mode"] = "transverse"
    transverse["secondary_cuts"] = [_solid(1.0, reinforcement=20.0, plate_uid="half")]
    tracks = [
        _track(_solid(90.0, reinforcement=40.0, plate_uid="recv")),
        _track(_solid(8.0, plate_uid="lead"), transverse),
    ]

    assert top_up_tracks_from_following(tracks).moves == 0
    assert tracks[1]["items"][-1]["plate_uid"] == "cut"


def test_split_without_secondary_cuts_is_not_a_donor():
    item = _solid(4.0, reinforcement=30.0, plate_uid="split")
    item["mode"] = "split"
    tracks = [
        _track(_solid(90.0, reinforcement=40.0, plate_uid="recv")),
        _track(_solid(8.0, plate_uid="lead"), item),
    ]

    assert top_up_tracks_from_following(tracks).moves == 0
    assert tracks[1]["items"][-1]["plate_uid"] == "split"


def test_receiver_continues_after_empty_donor_removed():
    tracks = [
        _track(_solid(91.0, reinforcement=40.0, concrete_grade="М500", plate_uid="recv")),
        _track(_solid(4.0, reinforcement=30.0, concrete_grade="М400", plate_uid="mid")),
        _track(_solid(3.0, reinforcement=30.0, concrete_grade="М400", plate_uid="last")),
    ]
    before_uids, before_len = _bundle(tracks)

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 2
    assert [item["plate_uid"] for item in tracks[0]["items"]] == ["recv", "last", "mid"]
    assert tracks[0]["length"] == pytest.approx(98.0, abs=0.01)
    _assert_conserved(before_uids, before_len, tracks)


def test_empty_donor_removed_when_gap_is_filled():
    tracks = [
        _track(_solid(97.0, reinforcement=40.0, plate_uid="recv")),
        _track(_solid(4.0, reinforcement=30.0, plate_uid="sole")),
    ]
    before_uids, before_len = _bundle(tracks)

    top_up_tracks_from_following(tracks)

    assert len(tracks) == 1
    assert tracks[0]["length"] == pytest.approx(101.0, abs=0.01)
    assert tracks[0]["items"][-1]["plate_uid"] == "sole"
    _assert_conserved(before_uids, before_len, tracks)


def _narrow_split(length: float, *, plate_uid: str, plate_name: str) -> dict:
    """Узкий продольный рез отдельным item: mode split, пустой secondary_cuts."""
    item = _solid(length, reinforcement=30.0, load_code=8, width=1.1, plate_uid=plate_uid)
    item["mode"] = "split"
    item["main_w"] = 1.1
    item["rest_w"] = 0.1
    item["plate_name"] = plate_name
    item["secondary_cuts"] = []
    return item


def test_last_solid_and_split_remainder_moves_in_order_and_cut_is_not_first():
    whole = _solid(1.8, reinforcement=30.0, load_code=8, width=1.2, plate_uid="whole")
    whole["plate_name"] = "ПБ 18-12-8"
    cut = _narrow_split(1.8, plate_uid="cut", plate_name="ПБ 18-11-8")
    tracks = [
        _track(_solid(91.14, reinforcement=30.0, load_code=8, plate_uid="body")),
        _track(whole, cut),
    ]
    before_uids, before_len = _bundle(tracks)

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 2
    assert len(tracks) == 1
    assert [item["plate_uid"] for item in tracks[0]["items"]] == ["body", "whole", "cut"]
    assert tracks[0]["items"][0]["mode"] == "solid"
    assert tracks[0]["items"][-1]["mode"] == "split"
    assert "11" in tracks[0]["items"][-1]["plate_name"]
    assert tracks[0]["length"] == pytest.approx(94.74, abs=0.01)
    _assert_conserved(before_uids, before_len, tracks)


def test_remainder_over_the_gap_stays_intact_while_non_first_plate_moves():
    # Сумма 4,7 м больше зазора 3,0 м. Целая 2,0 м и рез 1,5 м по отдельности
    # влезли бы и длиннее непервой плиты 1,2 м, но новым правилом хвост не рвём.
    cut = _narrow_split(1.5, plate_uid="cut", plate_name="ПБ 15-11-8")
    tracks = [
        _track(_solid(98.0, reinforcement=40.0, load_code=8, plate_uid="recv")),
        _track(
            _solid(2.0, reinforcement=30.0, load_code=8, width=1.2, plate_uid="head"),
            _solid(1.2, reinforcement=30.0, load_code=8, plate_uid="partial"),
            cut,
        ),
    ]
    before_uids, before_len = _bundle(tracks)

    result = top_up_tracks_from_following(tracks)

    assert result.moves == 1
    assert [item["plate_uid"] for item in tracks[0]["items"]] == ["recv", "partial"]
    assert tracks[0]["length"] == pytest.approx(99.2, abs=0.01)
    assert [item["plate_uid"] for item in tracks[1]["items"]] == ["head", "cut"]
    assert tracks[1]["items"][0]["mode"] == "solid"
    assert tracks[1]["items"][1]["mode"] == "split"
    assert tracks[1]["length"] == pytest.approx(3.5, abs=0.01)
    _assert_conserved(before_uids, before_len, tracks)
