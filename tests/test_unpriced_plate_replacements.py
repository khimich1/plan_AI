from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.unpriced_plate_replacements import (
    build_unpriced_plate_lines,
    list_lower_load_replacements,
    rewrite_plate_line_load,
)


def _create_price_db(path: Path) -> None:
    conn = sqlite3.connect(path)
    try:
        conn.execute(
            "CREATE TABLE prices (length_dm INTEGER, load_code INTEGER, price REAL, PRIMARY KEY(length_dm, load_code))"
        )
        conn.commit()
    finally:
        conn.close()


def _insert_price(path: Path, *, length_dm: int, load_code: int, price: float) -> None:
    conn = sqlite3.connect(path)
    try:
        conn.execute(
            "INSERT INTO prices (length_dm, load_code, price) VALUES (?, ?, ?)",
            (length_dm, load_code, price),
        )
        conn.commit()
    finally:
        conn.close()


def test_list_lower_load_replacements_for_75_12(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    _insert_price(db_path, length_dm=75, load_code=12, price=0.0)
    _insert_price(db_path, length_dm=75, load_code=10, price=31890.0)
    _insert_price(db_path, length_dm=75, load_code=8, price=29316.0)
    _insert_price(db_path, length_dm=75, load_code=6, price=27144.0)

    result = list_lower_load_replacements(length_dm=75, load_code=12, db_path=str(db_path))

    assert result == [
        {"load_code": 10, "price": 31890.0},
        {"load_code": 8, "price": 29316.0},
        {"load_code": 6, "price": 27144.0},
    ]


def test_list_lower_load_replacements_via_length_m(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    _insert_price(db_path, length_dm=75, load_code=10, price=100.0)
    _insert_price(db_path, length_dm=75, load_code=8, price=90.0)

    result = list_lower_load_replacements(length_m=7.5, load_code=12, db_path=str(db_path))
    assert [item["load_code"] for item in result] == [10, 8]


def test_list_lower_load_replacements_empty_when_no_prices(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    _insert_price(db_path, length_dm=75, load_code=12, price=0.0)
    _insert_price(db_path, length_dm=75, load_code=10, price=0.0)

    assert list_lower_load_replacements(length_dm=75, load_code=12, db_path=str(db_path)) == []


def test_list_lower_load_replacements_excludes_equal_or_higher(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    for code, price in ((12, 100.0), (10, 90.0), (8, 80.0)):
        _insert_price(db_path, length_dm=70, load_code=code, price=price)

    result = list_lower_load_replacements(length_dm=70, load_code=10, db_path=str(db_path))
    assert [item["load_code"] for item in result] == [8]
    assert 12 not in {item["load_code"] for item in result}
    assert 10 not in {item["load_code"] for item in result}


def test_rewrite_plate_line_load_handles_12_5() -> None:
    assert rewrite_plate_line_load("ПБ 75-12-12,5п 2", 10) == "ПБ 75-12-10п 2"
    assert rewrite_plate_line_load("Плиты ПБ 75-12-12п", 8) == "Плиты ПБ 75-12-8п"


def test_build_unpriced_plate_lines(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    _insert_price(db_path, length_dm=75, load_code=10, price=31890.0)
    _insert_price(db_path, length_dm=75, load_code=8, price=29316.0)

    lines = build_unpriced_plate_lines(
        [
            {
                "name": "Плиты ПБ 75-12-12п",
                "length_m": 7.5,
                "width_m": 1.2,
                "qty": 2,
                "load_class": 1200,
                "unit_price": None,
            },
            {
                "name": "Плиты ПБ 60-12-8п",
                "length_m": 6.0,
                "width_m": 1.2,
                "qty": 1,
                "load_class": 800,
                "unit_price": 25000.0,
            },
        ],
        db_path=str(db_path),
        normalized_lines=["ПБ 75-12-12п 2", "ПБ 60-12-8п 1"],
    )

    assert len(lines) == 1
    assert lines[0]["id"] == "unpriced-1"
    assert lines[0]["line"] == "ПБ 75-12-12п 2"
    assert lines[0]["qty"] == 2
    assert lines[0]["load_class"] == 1200
    assert [r["load_code"] for r in lines[0]["replacements"]] == [10, 8]
