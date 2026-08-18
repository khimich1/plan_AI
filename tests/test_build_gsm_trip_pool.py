"""Unit-тесты чистых функций сборки пула поездок ГСМ."""

from __future__ import annotations

from scripts.build_gsm_trip_pool import (
    TripFact,
    normalize_address,
    pair_rounds,
    parse_date_from_name,
)


def test_normalize_address_collapses_variants():
    a = normalize_address("г.Кострома, ул. Кузнецкая, д.18Б")
    b = normalize_address("г Кострома ул Кузнецкая дом 18Б")
    assert a == b
    assert "улица" not in a
    assert "дом" not in a


def test_normalize_address_strips_trailing_date():
    a = normalize_address("г.Кострома, ул. Кузнецкая, д.18Б, 02.04.2026г.")
    b = normalize_address("г.Кострома, ул. Кузнецкая, д.18Б, 03.04.2026г.")
    assert a == b
    assert "2026" not in a


def test_parse_date_from_name():
    assert parse_date_from_name("ПЛ 10.01.25.xls").isoformat() == "2025-01-10"
    assert parse_date_from_name("ПЛ 03.04.2025.xls").isoformat() == "2025-04-03"
    assert parse_date_from_name("Новый ПЛ 26.03.25.xls").isoformat() == "2025-03-26"


def test_pair_rounds_mirrors_neighbors():
    common = dict(
        vehicle_folder="Geely Monjaro",
        mark="Geely Monjaro",
        plate="О 165 ХУ 44",
        pl_date=None,
        driver="X",
        fuel="АИ-95",
        source_path="/tmp/a.xls",
        km=95.0,
    )
    t1 = TripFact(
        seq=1,
        addr_from="A",
        addr_to="B",
        addr_from_norm="a",
        addr_to_norm="b",
        time_dep="07:10",
        time_ret="09:00",
        **common,
    )
    t2 = TripFact(
        seq=2,
        addr_from="B",
        addr_to="A",
        addr_from_norm="b",
        addr_to_norm="a",
        time_dep="13:50",
        time_ret="16:00",
        **common,
    )
    t3 = TripFact(
        seq=1,
        addr_from="X",
        addr_to="Y",
        addr_from_norm="x",
        addr_to_norm="y",
        time_dep="10:00",
        time_ret="11:00",
        vehicle_folder="Geely Monjaro",
        mark="Geely Monjaro",
        plate="О 165 ХУ 44",
        pl_date=None,
        driver="X",
        fuel="АИ-95",
        source_path="/tmp/b.xls",
        km=10.0,
    )
    rounds, unpaired = pair_rounds([t1, t2, t3])
    assert len(rounds) == 1
    assert rounds[0].km_sum == 190.0
    assert len(unpaired) == 1
    assert unpaired[0].addr_from_norm == "x"
