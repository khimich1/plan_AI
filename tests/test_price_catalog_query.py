"""Unit tests for read-only factory price catalog listing."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.fbs_price_db import init_fbs_prices_schema
from core.march_price_db import init_march_prices_schema
from core.pile_price_db import init_pile_prices_schema
from core.price_catalog_query import CATALOG_CAP, list_price_catalog
from core.step_price_db import init_step_prices_schema


def _seed_pile_prices(db_path: Path, rows: list[tuple[str, str, float]]) -> None:
    init_pile_prices_schema(str(db_path))
    conn = sqlite3.connect(db_path)
    try:
        conn.executemany(
            "INSERT INTO pile_prices (mark, concrete_grade, price) VALUES (?, ?, ?)",
            rows,
        )
        conn.commit()
    finally:
        conn.close()


def _seed_step_prices(db_path: Path, rows: list[tuple[str, float]]) -> None:
    init_step_prices_schema(str(db_path))
    conn = sqlite3.connect(db_path)
    try:
        conn.executemany(
            "INSERT INTO step_prices (mark, price) VALUES (?, ?)",
            rows,
        )
        conn.commit()
    finally:
        conn.close()


def test_list_piles_q_c110_finds_neighbor_not_missing_mark(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _seed_pile_prices(
        db_path,
        [
            ("С110.30-9", "B25", 27585.43),
            ("С120.35-12", "B25", 44634.03),
            ("С110.30-9", "B20", 0.0),
        ],
    )

    items = list_price_catalog("piles", q="C110", db_path=str(db_path))
    marks = {(item["mark"], item["concrete_grade"]) for item in items}

    assert ("С110.30-9", "B25") in marks
    assert all(item["mark"] != "С110.30-6" for item in items)
    assert all(item["price"] > 0 for item in items)
    assert all(item["concrete_grade"] is not None for item in items)


def test_list_empty_q_on_small_fixture_does_not_raise(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _seed_pile_prices(db_path, [("С120.35-12", "B25", 44634.03)])

    items = list_price_catalog("piles", q="", db_path=str(db_path))
    assert len(items) == 1
    assert items[0]["mark"] == "С120.35-12"
    assert items[0]["price"] == pytest.approx(44634.03)


def test_list_steps_have_null_concrete_grade(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _seed_step_prices(db_path, [("ЛС11", 12000.0), ("ЛС14-1лев", 15000.0)])

    items = list_price_catalog("steps", q="ЛС11", db_path=str(db_path))
    assert len(items) == 1
    assert items[0]["mark"] == "ЛС11"
    assert items[0]["concrete_grade"] is None
    assert items[0]["price"] == pytest.approx(12000.0)


def test_unknown_product_type_raises_value_error(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    with pytest.raises(ValueError, match="Неизвестный"):
        list_price_catalog("widgets", q="", db_path=str(db_path))


def test_cap_exceeded_raises_value_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    db_path = tmp_path / "pb.db"
    _seed_pile_prices(
        db_path,
        [
            ("С110.30-9", "B25", 100.0),
            ("С110.30-10", "B25", 200.0),
            ("С110.35-9", "B25", 300.0),
        ],
    )
    monkeypatch.setattr("core.price_catalog_query.CATALOG_CAP", 2)

    with pytest.raises(ValueError, match="[Уу]точните поиск"):
        list_price_catalog("piles", q="C110", db_path=str(db_path))
    assert CATALOG_CAP == 2000


def test_missing_table_returns_empty_without_create(tmp_path: Path) -> None:
    db_path = tmp_path / "empty.db"
    sqlite3.connect(db_path).close()

    items = list_price_catalog("piles", q="C110", db_path=str(db_path))
    assert items == []

    conn = sqlite3.connect(db_path)
    try:
        tables = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
    finally:
        conn.close()
    assert "pile_prices" not in tables


def test_graded_types_init_without_write_on_empty_db(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    init_fbs_prices_schema(str(db_path))
    init_march_prices_schema(str(db_path))
    assert list_price_catalog("fbs", db_path=str(db_path)) == []
    assert list_price_catalog("marches", db_path=str(db_path)) == []
    assert list_price_catalog("bridge_piles", db_path=str(db_path)) == []
