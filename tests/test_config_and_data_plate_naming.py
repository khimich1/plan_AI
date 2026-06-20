"""Unit tests for plate naming and prays variant helpers in config_and_data."""

from __future__ import annotations

import pytest

from core.config_and_data import (
    extract_length_dm_raw_from_plate_name,
    format_reinforcement_from_load_code,
    make_plate_name,
    plate_name_to_prays_variant,
    plate_name_to_prays_variants,
)


@pytest.mark.parametrize(
    ("load_code", "expected"),
    [
        (8, "8п"),
        (10, "10п"),
        (12.5, "12,5п"),
        (0, "8п"),
        ("bad", "8п"),
    ],
)
def test_format_reinforcement_from_load_code(load_code, expected: str) -> None:
    assert format_reinforcement_from_load_code(load_code) == expected


@pytest.mark.parametrize(
    ("length_m", "width_m", "kwargs", "expected"),
    [
        (6.3, 1.2, {"load_code": 8}, "Плиты ПБ 63-12-8п"),
        (5.981, 1.2, {"load_code": 8, "length_dm_raw": "59,81"}, "Плиты ПБ 59,81-12-8п"),
        (4.0, 1.2, {"load_code": 8, "length_dm_raw": "40"}, "Плиты ПБ 40-12-8п"),
        (4.0, 1.2, {"load_code": 8, "length_dm_raw": "40,0"}, "Плиты ПБ 40,0-12-8п"),
        (4.2, 0.53, {"load_code": 10}, "Плиты ПБ 42-5,3-10п"),
        (2.5, 0.3, {"load_code": 8}, "Плиты ПБ 25-3-8п"),
        (2.5, 0.2, {"load_code": 6}, "Плиты ПБ 25-2-6п"),
    ],
)
def test_make_plate_name(length_m, width_m, kwargs, expected: str) -> None:
    assert make_plate_name(length_m, width_m, **kwargs) == expected


@pytest.mark.parametrize(
    ("name", "expected"),
    [
        ("Плиты ПБ 42-0.3-8п", "Плиты ПБ 42-3,0-8п"),
        ("Плиты ПБ 25-0.2-8п", "Плиты ПБ 25-2,0-8п"),
        ("Плиты ПБ 61,8-5-8п", "Плиты ПБ 61,8-5,0-8п"),
        ("Плиты ПБ 45-7-6п", "Плиты ПБ 45-7,0-6п"),
        ("Плиты ПБ 63-12-8п", None),
        ("Плиты ПБ 42-3,0-8п", None),
    ],
)
def test_plate_name_to_prays_variant(name: str, expected: str | None) -> None:
    assert plate_name_to_prays_variant(name) == expected


def test_plate_name_to_prays_variants_dedup_and_order() -> None:
    name = "Плиты ПБ 45-5-8п"
    variants = plate_name_to_prays_variants(name)

    assert variants
    assert name not in variants
    assert len(variants) == len(set(variants))
    assert "Плиты ПБ 45-5,0-8п" in variants
    assert "Плиты ПБ 45,0-5,0-8п" in variants


@pytest.mark.parametrize(
    ("plate_name", "expected_raw"),
    [
        ("Плиты ПБ 59,8-12-8п", "59,8"),
        ("ПБ 78-12-8п", "78"),
        ("Плиты ПБ 40,0-12-8п", "40,0"),
        ("garbage", None),
    ],
)
def test_extract_length_dm_raw_from_plate_name(plate_name: str, expected_raw: str | None) -> None:
    assert extract_length_dm_raw_from_plate_name(plate_name) == expected_raw
