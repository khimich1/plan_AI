"""Ручки себестоимости цельных свай в pb.db.

Миграция создаёт три таблицы и больше ничего. Первое чтение пустых таблиц
заполняет их числами 06.10.2026. Запись меняет только эти таблицы: pile_prices,
объём сваи и килограммы арматуры здесь не пишутся.
"""

from __future__ import annotations

import math
import re
import sqlite3
from collections.abc import Mapping
from dataclasses import dataclass

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
from core.pile_price_db import GRADE_CODES
from core.price_db import DEFAULT_DB, _connect

# Сырьё и расходы — таблица стартовых ручек в спеке.
# Четыре цены металла — руб/кг того же файла, уже зафиксированные формулой.
# Другие диаметры в спеке не названы, поэтому в старт не подставляются.
_DEFAULT_ARTICLE_PRICES: dict[str, float] = {
    ARTICLE_CEMENT: 7.54,
    ARTICLE_CRUSHED_20: 1.80,
    ARTICLE_CRUSHED_5: 1.97,
    ARTICLE_SAND: 0.66,
    ARTICLE_ADDITIVE: 59.0,
    ARTICLE_WIRE_5: 72.7,
    ARTICLE_AI_10: 50.57,
    ARTICLE_A500C_10: 51.39,
    ARTICLE_A500C_12: 51.39,
}

_DEFAULT_RECIPES: dict[str, Recipe] = {
    "B15": Recipe(cement_kg_m3=320, sand_kg_m3=990, crushed_kg_m3=950),
    "B20": Recipe(cement_kg_m3=340, sand_kg_m3=950, crushed_kg_m3=950),
    "B22_5": Recipe(cement_kg_m3=355, sand_kg_m3=940, crushed_kg_m3=950),
    "B25": Recipe(cement_kg_m3=370, sand_kg_m3=930, crushed_kg_m3=950),
    "B30_granite": Recipe(cement_kg_m3=400, sand_kg_m3=920, crushed_kg_m3=950),
}

_DEFAULT_OVERHEAD = Overhead(general=3186, payroll=3400, forming=40, other=606)

if tuple(_DEFAULT_RECIPES) != tuple(GRADE_CODES):
    raise RuntimeError("рецепт по умолчанию должен совпадать с GRADE_CODES")

_ARTICLE_PRICE = "pile_cost_article_price"
_RECIPE = "pile_cost_recipe"
_OVERHEAD = "pile_cost_overhead"
_TABLES = (_ARTICLE_PRICE, _RECIPE, _OVERHEAD)

_ARTICLE_CODE = re.compile(r"[a-z][a-z0-9_]*")

_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS pile_cost_article_price (
    article TEXT PRIMARY KEY,
    price REAL NOT NULL CHECK (price > 0)
);

CREATE TABLE IF NOT EXISTS pile_cost_recipe (
    concrete_grade TEXT PRIMARY KEY,
    cement_kg_m3 REAL NOT NULL CHECK (cement_kg_m3 >= 0),
    sand_kg_m3 REAL NOT NULL CHECK (sand_kg_m3 >= 0),
    crushed_kg_m3 REAL NOT NULL CHECK (crushed_kg_m3 >= 0)
);

CREATE TABLE IF NOT EXISTS pile_cost_overhead (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    general REAL NOT NULL CHECK (general >= 0),
    payroll REAL NOT NULL CHECK (payroll >= 0),
    forming REAL NOT NULL CHECK (forming >= 0),
    other REAL NOT NULL CHECK (other >= 0)
);
"""


@dataclass(frozen=True, slots=True)
class PileCostInputs:
    """Цены статей, рецепт по классу бетона и четыре статьи расходов."""

    prices: PriceBook
    recipes: Mapping[str, Recipe]
    overhead: Overhead

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "recipes",
            {str(grade): recipe for grade, recipe in self.recipes.items()},
        )


def ensure_pile_cost_inputs_schema(db_path: str = DEFAULT_DB) -> None:
    """Создаёт пустые таблицы ручек. Строки и pile_prices не пишет."""
    conn = _connect(db_path)
    try:
        conn.executescript(_SCHEMA_SQL)
        conn.commit()
    finally:
        conn.close()


def load_pile_cost_inputs(db_path: str = DEFAULT_DB) -> PileCostInputs:
    """Читает ручки. Пустые таблицы один раз заполняет числами 06.10.2026."""
    ensure_pile_cost_inputs_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        if _is_empty(conn):
            _write(conn, _defaults())
        loaded = _read(conn)
        conn.commit()
        return loaded
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def save_pile_cost_inputs(inputs: PileCostInputs, db_path: str = DEFAULT_DB) -> None:
    """Заменяет ручки целиком. Цена ≤ 0 и рецепт < 0 отклоняются до записи."""
    stored = _validated(inputs)
    ensure_pile_cost_inputs_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        _write(conn, stored)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _defaults() -> PileCostInputs:
    return _validated(
        PileCostInputs(
            prices=PriceBook(_DEFAULT_ARTICLE_PRICES),
            recipes=_DEFAULT_RECIPES,
            overhead=_DEFAULT_OVERHEAD,
        )
    )


def _validated(inputs: PileCostInputs) -> PileCostInputs:
    return PileCostInputs(
        prices=PriceBook(_validate_prices(inputs.prices.prices)),
        recipes=_validate_recipes(inputs.recipes),
        overhead=_validate_overhead(inputs.overhead),
    )


def _validate_prices(raw: Mapping[str, float]) -> dict[str, float]:
    cleaned: dict[str, float] = {}
    for article, value in raw.items():
        code = str(article).strip()
        if code in cleaned or not _ARTICLE_CODE.fullmatch(code):
            raise ValueError(f"неизвестный код статьи {article}")
        cleaned[code] = _positive(f"цена статьи {code}", value)
    missing = [code for code in _DEFAULT_ARTICLE_PRICES if code not in cleaned]
    if missing:
        raise ValueError("нет цены статьи " + ", ".join(missing))
    ordered = {code: cleaned[code] for code in _DEFAULT_ARTICLE_PRICES}
    for code in sorted(set(cleaned) - set(ordered)):
        ordered[code] = cleaned[code]
    return ordered


def _validate_recipes(raw: Mapping[str, Recipe]) -> dict[str, Recipe]:
    cleaned: dict[str, Recipe] = {}
    for grade, recipe in raw.items():
        code = str(grade).strip()
        if code not in GRADE_CODES or code in cleaned or not isinstance(recipe, Recipe):
            raise ValueError(f"класс {grade} не входит в рецепт")
        cleaned[code] = Recipe(
            cement_kg_m3=_non_negative(f"цемент рецепта {code}", recipe.cement_kg_m3),
            sand_kg_m3=_non_negative(f"песок рецепта {code}", recipe.sand_kg_m3),
            crushed_kg_m3=_non_negative(f"щебень рецепта {code}", recipe.crushed_kg_m3),
        )
    missing = [grade for grade in GRADE_CODES if grade not in cleaned]
    if missing:
        raise ValueError("нет рецепта класса " + ", ".join(missing))
    return {grade: cleaned[grade] for grade in GRADE_CODES}


def _validate_overhead(overhead: Overhead) -> Overhead:
    if not isinstance(overhead, Overhead):
        raise ValueError("расходы заданы неверно")
    return Overhead(
        general=_non_negative("общие производственные", overhead.general),
        payroll=_non_negative("ФОТ", overhead.payroll),
        forming=_non_negative("обслуживание формовки", overhead.forming),
        other=_non_negative("обслуживание, остальное", overhead.other),
    )


def _positive(label: str, value: object) -> float:
    number = _finite(label, value)
    if number <= 0:
        raise ValueError(f"{label} должна быть больше нуля")
    return number


def _non_negative(label: str, value: object) -> float:
    number = _finite(label, value)
    if number < 0:
        raise ValueError(f"{label} не может быть отрицательным")
    return number


def _finite(label: str, value: object) -> float:
    try:
        number = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} должно быть числом") from exc
    if not math.isfinite(number):
        raise ValueError(f"{label} должно быть конечным числом")
    return number


def _is_empty(conn: sqlite3.Connection) -> bool:
    return all(_count(conn, table) == 0 for table in _TABLES)


def _count(conn: sqlite3.Connection, table: str) -> int:
    row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    if row is None:
        return 0
    return int(row[0])


def _write(conn: sqlite3.Connection, inputs: PileCostInputs) -> None:
    """Заменяет строки трёх таблиц ручек. pile_prices не упоминается."""
    for table in _TABLES:
        conn.execute(f"DELETE FROM {table}")
    conn.executemany(
        f"INSERT INTO {_ARTICLE_PRICE} (article, price) VALUES (?, ?)",
        list(inputs.prices.prices.items()),
    )
    conn.executemany(
        f"""
        INSERT INTO {_RECIPE}
            (concrete_grade, cement_kg_m3, sand_kg_m3, crushed_kg_m3)
        VALUES (?, ?, ?, ?)
        """,
        [
            (
                grade,
                inputs.recipes[grade].cement_kg_m3,
                inputs.recipes[grade].sand_kg_m3,
                inputs.recipes[grade].crushed_kg_m3,
            )
            for grade in GRADE_CODES
        ],
    )
    overhead = inputs.overhead
    conn.execute(
        f"""
        INSERT INTO {_OVERHEAD} (id, general, payroll, forming, other)
        VALUES (1, ?, ?, ?, ?)
        """,
        (overhead.general, overhead.payroll, overhead.forming, overhead.other),
    )


def _read(conn: sqlite3.Connection) -> PileCostInputs:
    prices = {
        article: price
        for article, price in conn.execute(f"SELECT article, price FROM {_ARTICLE_PRICE}")
    }
    recipes = {
        grade: Recipe(
            cement_kg_m3=cement,
            sand_kg_m3=sand,
            crushed_kg_m3=crushed,
        )
        for grade, cement, sand, crushed in conn.execute(
            f"""
            SELECT concrete_grade, cement_kg_m3, sand_kg_m3, crushed_kg_m3
            FROM {_RECIPE}
            """
        )
    }
    row = conn.execute(
        f"SELECT general, payroll, forming, other FROM {_OVERHEAD} WHERE id = 1"
    ).fetchone()
    if row is None:
        raise ValueError("расходы себестоимости не заполнены")
    return _validated(
        PileCostInputs(
            prices=PriceBook(prices),
            recipes=recipes,
            overhead=Overhead(general=row[0], payroll=row[1], forming=row[2], other=row[3]),
        )
    )
