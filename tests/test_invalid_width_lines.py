from __future__ import annotations

import sqlite3
from pathlib import Path

from core.invalid_width_lines import build_invalid_width_lines


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


def _plate(
    *,
    name: str,
    length_m: float,
    width_m: float,
    qty: int = 1,
    load_class: int = 800,
    unit_price: float | None = 1000.0,
) -> dict:
    return {
        "name": name,
        "length_m": length_m,
        "width_m": width_m,
        "qty": qty,
        "load_class": load_class,
        "unit_price": unit_price,
        "product_type": "plates",
    }


def test_mix_12_and_8_one_invalid_with_neighbors(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)
    _insert_price(db_path, length_dm=29, load_code=8, price=10400.0)

    lines = build_invalid_width_lines(
        [
            _plate(name="Плиты ПБ 29-12-8п", length_m=2.9, width_m=1.2),
            _plate(name="Плиты ПБ 29-8-8п", length_m=2.9, width_m=0.8, qty=2),
        ],
        db_path=str(db_path),
        normalized_lines=["ПБ 29-12-8п 1", "ПБ 29-8-8п 2"],
    )

    assert len(lines) == 1
    assert lines[0]["id"] == "invalid-width-1"
    assert lines[0]["line"] == "ПБ 29-8-8п 2"
    assert lines[0]["width_mm"] == 800
    assert [r["width_mm"] for r in lines[0]["replacements"]] == [720, 860]
    assert [r["width_label"] for r in lines[0]["replacements"]] == ["7,2", "8,6"]
    assert all(r.get("price") == 10400.0 for r in lines[0]["replacements"])


def test_zero_three_and_three_not_invalid(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)

    lines = build_invalid_width_lines(
        [
            _plate(name="Плиты ПБ 78-0.3-8п", length_m=7.8, width_m=0.3),
            _plate(name="Плиты ПБ 78-3-8п", length_m=7.8, width_m=0.3),
        ],
        db_path=str(db_path),
        normalized_lines=["ПБ 78-0.3-8п 1", "ПБ 78-3-8п 1"],
    )
    assert lines == []


def test_wide_15_skipped_when_wide_lines_passed(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)

    lines = build_invalid_width_lines(
        [_plate(name="Плиты ПБ 60-15-8п", length_m=6.0, width_m=1.5)],
        db_path=str(db_path),
        normalized_lines=["ПБ 60-15-8п 1"],
        skip_wide_lines=["ПБ 60-15-8п 1"],
    )
    assert lines == []


def test_wide_15_skipped_by_width_even_without_skip_list(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)

    lines = build_invalid_width_lines(
        [_plate(name="Плиты ПБ 60-15-8п", length_m=6.0, width_m=1.5)],
        db_path=str(db_path),
        normalized_lines=["ПБ 60-15-8п 1"],
    )
    assert lines == []


def test_replacement_without_price_still_present(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)

    lines = build_invalid_width_lines(
        [_plate(name="Плиты ПБ 29-8-8п", length_m=2.9, width_m=0.8, unit_price=None)],
        db_path=str(db_path),
        normalized_lines=["ПБ 29-8-8п 1"],
    )

    assert len(lines) == 1
    assert [r["width_mm"] for r in lines[0]["replacements"]] == [720, 860]
    assert all("price" not in r or r["price"] is None for r in lines[0]["replacements"])


def test_non_plate_skipped(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _create_price_db(db_path)

    lines = build_invalid_width_lines(
        [
            {
                "name": "Свая С80.30-11",
                "length_m": 8.0,
                "width_m": 0.8,
                "qty": 1,
                "product_type": "piles",
                "product_kind": "pile",
            }
        ],
        db_path=str(db_path),
        normalized_lines=["Свая С80.30-11 1"],
    )
    assert lines == []
