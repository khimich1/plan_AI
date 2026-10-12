"""PC-002: снимок нормы цельных свай от 06.10.2026.

Ожидаемые цены — лист «Прайс», не подгонка под формулу.
xlsx в тесте не читается.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from core.pile_cost import (
    ARTICLE_A500C_10,
    ARTICLE_ADDITIVE,
    ARTICLE_AI_10,
    ARTICLE_CEMENT,
    ARTICLE_CRUSHED_20,
    ARTICLE_CRUSHED_5,
    ARTICLE_SAND,
    ARTICLE_WIRE_5,
    NormLine,
    Overhead,
    PriceBook,
    Recipe,
    buyer_price,
)

_NORM_PATH = Path(__file__).resolve().parents[1] / "core" / "pile_cost_norm.json"
_ROWS = json.loads(_NORM_PATH.read_text(encoding="utf-8"))["rows"]

_EXCLUDED_MARK = "С130.35-13А**"
_KEPT_STAR_MARK = "С130.35-13А*"

_PRICES = PriceBook(
    {
        ARTICLE_CEMENT: 7.54,
        ARTICLE_CRUSHED_20: 1.80,
        ARTICLE_CRUSHED_5: 1.97,
        ARTICLE_SAND: 0.66,
        ARTICLE_ADDITIVE: 59.0,
        ARTICLE_WIRE_5: 72.7,
        ARTICLE_AI_10: 50.57,
        ARTICLE_A500C_10: 51.39,
    }
)
_OVERHEAD = Overhead(general=3186, payroll=3400, forming=40, other=606)
_RECIPES = {
    "B15": Recipe(cement_kg_m3=320, sand_kg_m3=990, crushed_kg_m3=950),
    "B20": Recipe(cement_kg_m3=340, sand_kg_m3=950, crushed_kg_m3=950),
    "B22_5": Recipe(cement_kg_m3=355, sand_kg_m3=940, crushed_kg_m3=950),
    "B25": Recipe(cement_kg_m3=370, sand_kg_m3=930, crushed_kg_m3=950),
    "B30_granite": Recipe(cement_kg_m3=400, sand_kg_m3=920, crushed_kg_m3=950),
}
_S30_BUYER = {
    "B15": 3513.2818728,
    "B20": 3538.3051464,
    "B22_5": 3559.5848256,
    "B25": 3580.8645048,
    "B30_granite": 3655.4166552,
}
_TOLERANCE = 0.01


def _row(mark: str, grade: str) -> dict[str, object]:
    found = [row for row in _ROWS if row["mark"] == mark and row["concrete_grade"] == grade]
    assert len(found) == 1
    return found[0]


def _kg(row: dict[str, object], article: str) -> float:
    lines = [line for line in row["lines"] if line["article"] == article]  # type: ignore[index]
    assert len(lines) == 1
    return float(lines[0]["kg"])


def _buyer(row: dict[str, object]) -> float:
    grade = str(row["concrete_grade"])
    lines = tuple(NormLine(str(line["article"]), float(line["kg"])) for line in row["lines"])  # type: ignore[union-attr]
    additive = row.get("additive_kg")
    return buyer_price(
        section_cm=int(row["section_cm"]),  # type: ignore[arg-type]
        volume_m3=float(row["volume_m3"]),  # type: ignore[arg-type]
        concrete_grade=grade,
        recipe=_RECIPES[grade],
        lines=lines,
        prices=_PRICES,
        overhead=_OVERHEAD,
        additive_kg=None if additive is None else float(additive),  # type: ignore[arg-type]
    ).buyer_price


def test_snapshot_covers_price_sheet_marks() -> None:
    marks = {row["mark"] for row in _ROWS}
    assert len(_ROWS) == 1585
    assert len(marks) == 317
    assert all(row["concrete_grade"] in _RECIPES for row in _ROWS)


def test_double_asterisk_mark_is_absent() -> None:
    marks = {row["mark"] for row in _ROWS}
    assert _EXCLUDED_MARK not in marks
    assert all("**" not in mark for mark in marks)


def test_single_asterisk_mark_is_present() -> None:
    grades = {row["concrete_grade"] for row in _ROWS if row["mark"] == _KEPT_STAR_MARK}
    assert grades == set(_RECIPES)


def test_s30_20_3_b15_stores_volume_and_priced_metal() -> None:
    row = _row("С30.20-3", "B15")
    assert row["section_cm"] == 20
    assert row["volume_m3"] == pytest.approx(0.13)
    assert "additive_kg" not in row
    assert _kg(row, ARTICLE_WIRE_5) == pytest.approx(3.1)
    assert _kg(row, ARTICLE_AI_10) == pytest.approx(3.15)
    assert _kg(row, ARTICLE_A500C_10) == pytest.approx(8)
    assert {line["article"] for line in row["lines"]} == {  # type: ignore[index]
        ARTICLE_WIRE_5,
        ARTICLE_AI_10,
        ARTICLE_A500C_10,
    }


@pytest.mark.parametrize("grade", ["B15", "B20", "B22_5", "B25", "B30_granite"])
def test_s30_20_3_matches_price_sheet(grade: str) -> None:
    assert _buyer(_row("С30.20-3", grade)) == pytest.approx(_S30_BUYER[grade], abs=_TOLERANCE)


def test_s40_30_1_b15_matches_price_sheet() -> None:
    row = _row("С40.30-1", "B15")
    assert row["section_cm"] == 30
    assert row["volume_m3"] == pytest.approx(0.37)
    assert _kg(row, ARTICLE_WIRE_5) == pytest.approx(6.5)
    assert _kg(row, ARTICLE_AI_10) == pytest.approx(20)
    assert {line["article"] for line in row["lines"]} == {ARTICLE_WIRE_5, ARTICLE_AI_10}  # type: ignore[index]
    assert _buyer(row) == pytest.approx(9033.108758784, abs=_TOLERANCE)


def test_hand_entered_additive_is_stored_and_default_formula_is_not() -> None:
    typed = _row("С90.30-9", "B15")
    assert typed["additive_kg"] == pytest.approx(4.6)
    assert "additive_kg" not in _row("С30.20-3", "B20")
