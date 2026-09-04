from __future__ import annotations

import pytest

from core.factory_width import (
    format_factory_width_label,
    is_factory_width_mm,
    rewrite_plate_line_width,
    suggest_factory_width_mm,
    width_m_to_mm,
)


@pytest.mark.parametrize(
    ("width_mm", "expected"),
    [
        (260, True),
        (300, True),
        (320, True),
        (460, True),
        (530, True),
        (660, True),
        (720, True),
        (860, True),
        (920, True),
        (1020, True),
        (1080, True),
        (1200, True),
        (200, False),
        (400, False),
        (600, False),
        (800, False),
        (1000, False),
        (1100, False),
        (1190, False),
        (1500, False),
    ],
)
def test_is_factory_width_mm(width_mm: int, expected: bool) -> None:
    assert is_factory_width_mm(width_mm) is expected


@pytest.mark.parametrize(
    ("width_mm", "expected"),
    [
        (800, [720, 860]),
        (1000, [920, 1020]),
        (200, [260]),
        (1100, [1080, 1200]),
        (400, [320, 460]),
        (300, []),
        (720, []),
        (1200, []),
    ],
)
def test_suggest_factory_width_mm(width_mm: int, expected: list[int]) -> None:
    assert suggest_factory_width_mm(width_mm) == expected


@pytest.mark.parametrize(
    ("width_mm", "label"),
    [
        (720, "7,2"),
        (860, "8,6"),
        (1200, "12"),
        (260, "2,6"),
        (1020, "10,2"),
        (1080, "10,8"),
        (920, "9,2"),
    ],
)
def test_format_factory_width_label(width_mm: int, label: str) -> None:
    assert format_factory_width_label(width_mm) == label


def test_width_m_to_mm_rounds_parsed_width() -> None:
    assert width_m_to_mm(0.8) == 800
    assert width_m_to_mm(0.3) == 300
    assert width_m_to_mm(0.2) == 200
    assert width_m_to_mm(1.2) == 1200
    assert width_m_to_mm(1.0) == 1000


def test_rewrite_plate_line_width_keeps_qty_and_load() -> None:
    assert rewrite_plate_line_width("Плиты ПБ 29-8-8п", 860) == "Плиты ПБ 29-8,6-8п"
    assert rewrite_plate_line_width("ПБ 29-8-8п 2", 720) == "ПБ 29-7,2-8п 2"
    assert rewrite_plate_line_width("Плиты ПБ 60-10-8п", 1020) == "Плиты ПБ 60-10,2-8п"


def test_rewrite_plate_line_width_compact_without_pe_suffix() -> None:
    assert rewrite_plate_line_width("68-11-10 1", 1080) == "68-10,8-10 1"
    assert rewrite_plate_line_width("68-11-10", 1080) == "68-10,8-10"
    assert rewrite_plate_line_width("ПБ 68-11-10 1", 1200) == "ПБ 68-12-10 1"


def test_rewrite_plate_line_width_does_not_touch_zero_three_unless_called() -> None:
    assert rewrite_plate_line_width("ПБ 78-0.3-8п 1", 260) == "ПБ 78-2,6-8п 1"
