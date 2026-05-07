#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Unit-тесты FFD укладки кусков в дорожки (core.optimization.ffd_packing)."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.optimization.ffd_packing import Piece, Track, first_fit_decreasing, pack_tracks  # noqa: E402


def test_first_fit_decreasing_is_pack_tracks():
    """Alias для совместимости имён совпадает с основной функцией."""
    assert first_fit_decreasing is pack_tracks


def test_pack_tracks_two_pieces_one_track_when_stock_allows():
    """Два коротких куска помещаются в одну заготовку."""
    pieces = [
        Piece(length_m=5.0, qty=1, kind="standard", load_class=8.0),
        Piece(length_m=4.0, qty=1, kind="addon", load_class=8.0),
    ]
    tracks = pack_tracks(pieces, stock_len_m=10.0)
    assert len(tracks) == 1
    assert tracks[0].total_m == pytest.approx(9.0)
    assert tracks[0].leftover_m == pytest.approx(1.0)


def test_pack_tracks_two_pieces_force_two_tracks_when_sum_exceeds_stock():
    """Длинные куски не суммируются на одну плиту — две дорожки."""
    pieces = [
        Piece(length_m=6.0, qty=1, kind="standard", load_class=8.0),
        Piece(length_m=6.0, qty=1, kind="standard", load_class=8.0),
    ]
    tracks = pack_tracks(pieces, stock_len_m=10.0)
    assert len(tracks) == 2
    assert all(t.total_m == pytest.approx(6.0) for t in tracks)
    assert all(t.leftover_m == pytest.approx(4.0) for t in tracks)


def test_pack_tracks_three_pieces_two_tracks_ffd_placement():
    """
    Три куска: два длинные (7 м) разносятся по дорожкам,
    короткий (3 м) добавляется в первую дорожку (FFD по убыванию длины).
    """
    pieces = [
        Piece(length_m=3.0, qty=1, kind="standard", load_class=8.0),
        Piece(length_m=7.0, qty=2, kind="standard", load_class=8.0),
    ]
    tracks = pack_tracks(pieces, stock_len_m=10.0)
    assert len(tracks) == 2
    totals_sorted = sorted(t.total_m for t in tracks)
    assert totals_sorted[0] == pytest.approx(7.0)
    assert totals_sorted[1] == pytest.approx(10.0)


def test_track_pieces_annotation_is_list_of_piece():
    """При postponed annotations строка задаёт связь дорожки с Piece."""
    assert Track.__annotations__["pieces"] == "list[Piece]"


def test_pack_tracks_track_pieces_are_piece_instances_and_qty_expanded():
    """После раскладывания каждая единица — отдельный Piece с qty=1."""
    pieces = [Piece(length_m=4.0, qty=3, kind="standard", load_class=8.0)]
    tracks = pack_tracks(pieces, stock_len_m=10.0)
    assert len(tracks) == 2
    joined: list[Piece] = []
    for t in tracks:
        joined.extend(t.pieces)
    assert len(joined) == 3
    for p in joined:
        assert isinstance(p, Piece)
        assert p.qty == 1
        assert p.length_m == pytest.approx(4.0)
