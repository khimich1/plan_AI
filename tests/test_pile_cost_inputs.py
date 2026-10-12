"""PC-003: ручки себестоимости в pb.db, без записи pile_prices."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.pile_cost import (
    ARTICLE_A500C_10,
    ARTICLE_A500C_12,
    ARTICLE_ADDITIVE,
    ARTICLE_AI_10,
    ARTICLE_CEMENT,
    ARTICLE_CRUSHED_20,
    ARTICLE_CRUSHED_5,
    ARTICLE_SAND,
    ARTICLE_WIRE_5,
    Overhead,
    PriceBook,
    Recipe,
)
from core.pile_cost_inputs import (
    PileCostInputs,
    ensure_pile_cost_inputs_schema,
    load_pile_cost_inputs,
    save_pile_cost_inputs,
)
from core.pile_price_db import GRADE_CODES, init_pile_prices_schema

_SPEC_RECIPES = {
    "B15": Recipe(cement_kg_m3=320, sand_kg_m3=990, crushed_kg_m3=950),
    "B20": Recipe(cement_kg_m3=340, sand_kg_m3=950, crushed_kg_m3=950),
    "B22_5": Recipe(cement_kg_m3=355, sand_kg_m3=940, crushed_kg_m3=950),
    "B25": Recipe(cement_kg_m3=370, sand_kg_m3=930, crushed_kg_m3=950),
    "B30_granite": Recipe(cement_kg_m3=400, sand_kg_m3=920, crushed_kg_m3=950),
}

_INPUT_TABLES = (
    "pile_cost_article_price",
    "pile_cost_recipe",
    "pile_cost_overhead",
)


def _db(tmp_path: Path) -> str:
    return str(tmp_path / "pb.db")


def _with_price(inputs: PileCostInputs, article: str, price: float) -> PileCostInputs:
    prices = dict(inputs.prices.prices)
    prices[article] = price
    return PileCostInputs(prices=PriceBook(prices), recipes=inputs.recipes, overhead=inputs.overhead)


def _pile_snapshot(db_path: str) -> tuple[str | None, list[tuple]]:
    conn = sqlite3.connect(db_path)
    try:
        schema = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='pile_prices'"
        ).fetchone()
        if schema is None:
            return None, []
        rows = conn.execute(
            """
            SELECT mark, concrete_grade, price, price_list_date, imported_at
            FROM pile_prices
            ORDER BY mark, concrete_grade
            """
        ).fetchall()
        return schema[0], rows
    finally:
        conn.close()


def _seed_pile_prices(db_path: str) -> None:
    init_pile_prices_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.executemany(
            """
            INSERT INTO pile_prices
                (mark, concrete_grade, price, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            [
                ("С30.20-3", "B15", 3513.2818728, "2026-10-06", "2026-10-06T12:00:00"),
                ("С130.35-13А**", "B15", 60980.85, "2026-10-06", "2026-10-06T12:00:00"),
            ],
        )
        conn.commit()
    finally:
        conn.close()


def test_migration_creates_empty_input_tables_and_skips_pile_prices(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    ensure_pile_cost_inputs_schema(db_path)
    ensure_pile_cost_inputs_schema(db_path)

    conn = sqlite3.connect(db_path)
    try:
        tables = {
            row[0]
            for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        }
        assert set(_INPUT_TABLES) <= tables
        assert "pile_prices" not in tables
        for table in _INPUT_TABLES:
            assert conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0] == 0
            columns = [row[1] for row in conn.execute(f"PRAGMA table_info({table})")]
            assert "volume_m3" not in columns
            assert "rebar_kg" not in columns
            assert "kg" not in columns
    finally:
        conn.close()


def test_first_read_seeds_spec_handles(tmp_path: Path) -> None:
    loaded = load_pile_cost_inputs(_db(tmp_path))

    assert loaded.prices.unit_price(ARTICLE_CEMENT) == 7.54
    assert loaded.prices.unit_price(ARTICLE_CRUSHED_20) == 1.80
    assert loaded.prices.unit_price(ARTICLE_CRUSHED_5) == 1.97
    assert loaded.prices.unit_price(ARTICLE_SAND) == 0.66
    assert loaded.prices.unit_price(ARTICLE_ADDITIVE) == 59
    assert loaded.prices.unit_price(ARTICLE_WIRE_5) == 72.7
    assert loaded.prices.unit_price(ARTICLE_AI_10) == 50.57
    assert loaded.prices.unit_price(ARTICLE_A500C_10) == 51.39
    assert loaded.prices.unit_price(ARTICLE_A500C_12) == 51.39
    assert loaded.overhead == Overhead(general=3186, payroll=3400, forming=40, other=606)
    assert set(loaded.recipes) == set(GRADE_CODES)
    for grade in GRADE_CODES:
        assert loaded.recipes[grade] == _SPEC_RECIPES[grade]


def test_reread_returns_saved_cement_without_reset(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    loaded = load_pile_cost_inputs(db_path)
    save_pile_cost_inputs(_with_price(loaded, ARTICLE_CEMENT, 8.25), db_path)

    again = load_pile_cost_inputs(db_path)
    third = load_pile_cost_inputs(db_path)

    assert again.prices.unit_price(ARTICLE_CEMENT) == 8.25
    assert third.prices.unit_price(ARTICLE_CEMENT) == 8.25
    assert again.prices.unit_price(ARTICLE_SAND) == 0.66
    assert again.recipes["B15"] == _SPEC_RECIPES["B15"]
    assert again.overhead.general == 3186


def test_save_does_not_change_pile_prices(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    _seed_pile_prices(db_path)
    before = _pile_snapshot(db_path)

    loaded = load_pile_cost_inputs(db_path)
    save_pile_cost_inputs(_with_price(loaded, ARTICLE_CEMENT, 9.5), db_path)

    assert _pile_snapshot(db_path) == before
    assert load_pile_cost_inputs(db_path).prices.unit_price(ARTICLE_CEMENT) == 9.5


def test_price_is_stored_without_rounding_to_the_ruble(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    loaded = load_pile_cost_inputs(db_path)
    fractional = 7.54123456789
    save_pile_cost_inputs(_with_price(loaded, ARTICLE_CEMENT, fractional), db_path)

    assert load_pile_cost_inputs(db_path).prices.unit_price(ARTICLE_CEMENT) == fractional


def test_extra_article_price_roundtrips(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    loaded = load_pile_cost_inputs(db_path)
    prices = dict(loaded.prices.prices)
    prices["probe_1"] = 48.5
    save_pile_cost_inputs(
        PileCostInputs(prices=PriceBook(prices), recipes=loaded.recipes, overhead=loaded.overhead),
        db_path,
    )

    assert load_pile_cost_inputs(db_path).prices.unit_price("probe_1") == 48.5


@pytest.mark.parametrize("price", [0, -1.5])
def test_non_positive_price_is_rejected_and_store_stays(tmp_path: Path, price: float) -> None:
    db_path = _db(tmp_path)
    _seed_pile_prices(db_path)
    before = _pile_snapshot(db_path)
    loaded = load_pile_cost_inputs(db_path)
    save_pile_cost_inputs(_with_price(loaded, ARTICLE_CEMENT, 8.25), db_path)

    with pytest.raises(ValueError):
        save_pile_cost_inputs(_with_price(loaded, ARTICLE_CEMENT, price), db_path)

    assert load_pile_cost_inputs(db_path).prices.unit_price(ARTICLE_CEMENT) == 8.25
    assert _pile_snapshot(db_path) == before


def test_negative_recipe_is_rejected_and_store_stays(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    loaded = load_pile_cost_inputs(db_path)
    recipes = dict(loaded.recipes)
    recipes["B15"] = Recipe(cement_kg_m3=-1, sand_kg_m3=990, crushed_kg_m3=950)

    with pytest.raises(ValueError):
        save_pile_cost_inputs(
            PileCostInputs(prices=loaded.prices, recipes=recipes, overhead=loaded.overhead),
            db_path,
        )

    assert load_pile_cost_inputs(db_path).recipes["B15"] == _SPEC_RECIPES["B15"]


def test_zero_recipe_and_zero_overhead_are_stored(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    loaded = load_pile_cost_inputs(db_path)
    recipes = dict(loaded.recipes)
    recipes["B20"] = Recipe(cement_kg_m3=0, sand_kg_m3=0, crushed_kg_m3=0)
    overhead = Overhead(general=0, payroll=0, forming=0, other=0)
    save_pile_cost_inputs(
        PileCostInputs(prices=loaded.prices, recipes=recipes, overhead=overhead),
        db_path,
    )

    again = load_pile_cost_inputs(db_path)
    assert again.recipes["B20"] == Recipe(cement_kg_m3=0, sand_kg_m3=0, crushed_kg_m3=0)
    assert again.overhead == overhead


def test_negative_overhead_is_rejected(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    loaded = load_pile_cost_inputs(db_path)

    with pytest.raises(ValueError):
        save_pile_cost_inputs(
            PileCostInputs(
                prices=loaded.prices,
                recipes=loaded.recipes,
                overhead=Overhead(general=-1, payroll=3400, forming=40, other=606),
            ),
            db_path,
        )

    assert load_pile_cost_inputs(db_path).overhead.general == 3186


def test_incomplete_recipe_is_rejected(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    loaded = load_pile_cost_inputs(db_path)
    recipes = dict(loaded.recipes)
    del recipes["B25"]

    with pytest.raises(ValueError):
        save_pile_cost_inputs(
            PileCostInputs(prices=loaded.prices, recipes=recipes, overhead=loaded.overhead),
            db_path,
        )

    assert load_pile_cost_inputs(db_path).recipes["B25"].cement_kg_m3 == 370


def test_schema_check_rejects_non_positive_price_and_negative_recipe(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    ensure_pile_cost_inputs_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                "INSERT INTO pile_cost_article_price (article, price) VALUES ('cement', 0)"
            )
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                """
                INSERT INTO pile_cost_recipe
                    (concrete_grade, cement_kg_m3, sand_kg_m3, crushed_kg_m3)
                VALUES ('B15', -1, 0, 0)
                """
            )
    finally:
        conn.close()
