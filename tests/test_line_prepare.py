"""Unit tests for core.line_prepare (shared cleanup + bridge GOST rewrite)."""

from __future__ import annotations

import pytest

from core.bridge_pile_line_parser import parse_bridge_pile_line
from core.line_prepare import cleanup_source_line, prepare_bridge_pile_mark, prepare_source_line
from core.pile_line_parser import parse_pile_line
from core.plate_line_parser import parse_line


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("C14-35T7 80шт", "C14-35T7 80"),
        ("C14-35T7 80 шт", "C14-35T7 80"),
        ("C14-35T7 80 шт.", "C14-35T7 80"),
        ("C14-35T7 80 штук", "C14-35T7 80"),
        ("1. C14-35T7 80", "C14-35T7 80"),
        ("1) C14-35T7 80", "C14-35T7 80"),
        ("1, C14-35T7 80", "C14-35T7 80"),
        ("C14-35T7 80", "C14-35T7 80"),
        ("ПБ 78-12-8п 5шт", "ПБ 78-12-8п 5"),
        ("С120.35-12 5шт", "С120.35-12 5"),
    ],
)
def test_shared_cleanup_qty_and_list_prefix(raw: str, expected: str) -> None:
    assert prepare_source_line(raw, "plates") == expected
    assert cleanup_source_line(raw) == expected


def test_nbsp_and_unicode_dashes() -> None:
    raw = "C14-35T7\u00a080\u00a0шт"
    assert prepare_source_line(raw, "bridge_piles") == "C14-35T7 80"
    dashed = "C14–35T7 80"
    assert "-" in prepare_source_line(dashed, "bridge_piles")
    assert "–" not in prepare_source_line(dashed, "bridge_piles")


def test_sht_in_middle_of_mark_is_not_stripped() -> None:
    raw = "C14шт35T7 80"
    prepared = prepare_source_line(raw, "bridge_piles")
    assert "шт" in prepared
    assert prepared.endswith("80")


def test_canonical_line_is_idempotent() -> None:
    clean = "C14-35T7 80"
    assert prepare_source_line(clean, "bridge_piles") == clean
    assert prepare_source_line(prepare_source_line(clean, "bridge_piles"), "bridge_piles") == clean


@pytest.mark.parametrize(
    "raw",
    [
        "C 14.35-T7 80",
        "C 14.35-T7 80 шт",
        "C14.35-T7 80шт",
    ],
)
def test_gost_rewrite_is_parseable_for_bridge(raw: str) -> None:
    prepared = prepare_source_line(raw, "bridge_piles")
    result = parse_bridge_pile_line(prepared)
    assert result.parsed is True
    assert result.qty == 80
    assert prepare_source_line(prepared, "bridge_piles") == prepared


def test_gost_without_class_is_not_rewritten() -> None:
    raw = "C 14.35 80"
    prepared = prepare_source_line(raw, "bridge_piles")
    assert "14.35" in prepared.replace(" ", "")
    assert parse_bridge_pile_line(prepared).parsed is False


def test_factory_bridge_mark_keeps_meaning() -> None:
    assert prepare_source_line("C14-35T7 80", "bridge_piles") == "C14-35T7 80"
    assert parse_bridge_pile_line(prepare_source_line("C14-35T7 80", "bridge_piles")).parsed


def test_prepare_bridge_pile_mark_gost_and_factory() -> None:
    assert prepare_bridge_pile_mark("C 14.35-T7") == "C14-35T7"
    assert prepare_bridge_pile_mark("C14-35T7") == "C14-35T7"
    assert prepare_bridge_pile_mark("C 14.35") == "C 14.35"
    rewritten = prepare_bridge_pile_mark("С 8.35-Т1")
    assert rewritten == "С8-35Т1"


def test_gost_rewrite_not_applied_for_other_types() -> None:
    raw = "C 14.35-T7 80 шт"
    plates = prepare_source_line(raw, "plates")
    assert "14.35" in plates
    assert parse_line(plates).parsed is False


def test_prepare_then_parse_plate_and_pile() -> None:
    plate = parse_line(prepare_source_line("ПБ 78-12-8п 5шт", "plates"))
    assert plate.parsed is True
    pile = parse_pile_line(prepare_source_line("С120.35-12 5шт", "piles"))
    assert pile.parsed is True
    assert pile.qty == 5
