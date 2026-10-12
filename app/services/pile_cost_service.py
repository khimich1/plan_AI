"""Превью и запись цен цельных свай.

Превью считает дифф и не пишет pile_prices. Apply пишет INSERT OR REPLACE
только ключей нормы, и только если inputs_sha256 совпал с сохранёнными ручками.
Объём сваи и килограммы арматуры здесь не пишутся.
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.core.settings import get_settings
from app.schemas.pile_cost import (
    PileCostArticlePrice,
    PileCostDiffExample,
    PileCostDiffExamples,
    PileCostGradeBreakdown,
    PileCostGradePrice,
    PileCostInputsBody,
    PileCostMarkBreakdown,
    PileCostMarkSummary,
    PileCostOverheadInput,
    PileCostPreview,
    PileCostPricedLine,
    PileCostRecipeInput,
    PileCostSaved,
    PileCostState,
)
from core.pile_cost import (
    NormLine,
    Overhead,
    PileCost,
    PriceBook,
    PricedLine,
    Recipe,
    buyer_price,
)
from core.pile_cost_inputs import PileCostInputs, load_pile_cost_inputs, save_pile_cost_inputs
from core.pile_price_db import GRADE_CODES, init_pile_prices_schema
from core.price_db import _connect
from core.price_desk_diff import DiffExample, PriceDeskDiffExamples, diff_price_maps

MSG_STALE = "расчёт изменился, посмотрите дифф ещё раз"
MSG_NEED_SHA = "укажите inputs_sha256 с превью"
MSG_MARK_NOT_FOUND = "марка не найдена"

_NORM_PATH = Path(__file__).resolve().parents[2] / "core" / "pile_cost_norm.json"
_PriceKey = tuple[str, str]


class PileCostConflict(Exception):
    """Отпечаток ручек не совпал с превью. pile_prices не меняется."""

    def __init__(self, message: str, status_code: int = 409) -> None:
        super().__init__(message)
        self.message = message
        self.status_code = status_code


class PileCostNotFound(Exception):
    """Марки нет в снимке нормы."""

    def __init__(self, message: str = MSG_MARK_NOT_FOUND, status_code: int = 404) -> None:
        super().__init__(message)
        self.message = message
        self.status_code = status_code


@dataclass(frozen=True, slots=True)
class _NormRow:
    mark: str
    concrete_grade: str
    section_cm: int
    volume_m3: float
    lines: tuple[NormLine, ...]
    additive_kg: float | None


@dataclass(frozen=True, slots=True)
class _MarkNorm:
    mark: str
    section_cm: int
    volume_m3: float
    grades: Mapping[str, _NormRow]


def inputs_sha256(inputs: PileCostInputs) -> str:
    """Отпечаток сохранённых ручек. Норма, объём и килограммы в него не входят."""
    raw = json.dumps(
        _fingerprint_payload(inputs),
        ensure_ascii=False,
        separators=(",", ":"),
        allow_nan=False,
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


class PileCostService:
    def __init__(self, db_path: Path | str | None = None) -> None:
        self._db_path = None if db_path is None else Path(db_path)

    def db_path(self) -> str:
        if self._db_path is not None:
            return str(self._db_path)
        return str(get_settings().pb_db_path)

    def view(self, mark: str | None = None) -> PileCostState:
        """Ручки, краткий список марок и разложение одной марки. pile_prices не пишет."""
        inputs = load_pile_cost_inputs(self.db_path())
        marks = _mark_index(_load_norm_rows())
        wanted = (mark or "").strip()
        breakdown = _breakdown(wanted, marks, inputs) if wanted else None
        return PileCostState(
            inputs_sha256=inputs_sha256(inputs),
            inputs=_inputs_body(inputs),
            marks=[_summary(item, inputs) for item in marks],
            breakdown=breakdown,
        )

    def save(self, body: PileCostInputsBody) -> PileCostSaved:
        """Сохраняет ручки. pile_prices, объём и килограммы арматуры не пишет."""
        save_pile_cost_inputs(_inputs_from_body(body), self.db_path())
        stored = load_pile_cost_inputs(self.db_path())
        return PileCostSaved(inputs_sha256=inputs_sha256(stored), inputs=_inputs_body(stored))

    def preview(self) -> PileCostPreview:
        """Дифф формулы и pile_prices. Таблицу цен не пишет."""
        db_path = self.db_path()
        inputs = load_pile_cost_inputs(db_path)
        quoted = _quoted_prices(inputs)
        return _preview(inputs, quoted, _current_prices(db_path))

    def apply(self, inputs_sha256_hex: str | None) -> PileCostPreview:
        """Пишет цены нормы, если отпечаток совпал. Иначе цены не меняются."""
        expected = _require_sha(inputs_sha256_hex)
        db_path = self.db_path()
        inputs = load_pile_cost_inputs(db_path)
        if inputs_sha256(inputs) != expected:
            raise PileCostConflict(MSG_STALE)
        quoted = _quoted_prices(inputs)
        current = _current_prices(db_path)
        _write_norm_prices(db_path, quoted, written_at=datetime.now())
        return _preview(inputs, quoted, current)


def _require_sha(value: str | None) -> str:
    text = (value or "").strip().lower()
    if not text:
        raise PileCostConflict(MSG_NEED_SHA)
    return text


def _fingerprint_payload(inputs: PileCostInputs) -> dict[str, object]:
    overhead = inputs.overhead
    return {
        "overhead": {
            "forming": overhead.forming,
            "general": overhead.general,
            "other": overhead.other,
            "payroll": overhead.payroll,
        },
        "prices": [
            {"article": article, "price": inputs.prices.prices[article]}
            for article in sorted(inputs.prices.prices)
        ],
        "recipes": [_recipe_payload(inputs, grade) for grade in GRADE_CODES],
    }


def _recipe_payload(inputs: PileCostInputs, grade: str) -> dict[str, object]:
    recipe = inputs.recipes[grade]
    return {
        "cement_kg_m3": recipe.cement_kg_m3,
        "concrete_grade": grade,
        "crushed_kg_m3": recipe.crushed_kg_m3,
        "sand_kg_m3": recipe.sand_kg_m3,
    }


def _load_norm_rows() -> tuple[_NormRow, ...]:
    payload = json.loads(_NORM_PATH.read_text(encoding="utf-8"))
    return tuple(_norm_row(raw) for raw in payload["rows"])


def _norm_row(raw: Mapping[str, object]) -> _NormRow:
    lines = raw["lines"]
    if not isinstance(lines, list):
        raise ValueError("строка нормы без списка металла")
    additive = raw.get("additive_kg")
    return _NormRow(
        mark=str(raw["mark"]),
        concrete_grade=str(raw["concrete_grade"]),
        section_cm=int(_number(raw["section_cm"])),
        volume_m3=_number(raw["volume_m3"]),
        lines=tuple(_norm_line(line) for line in lines),
        additive_kg=None if additive is None else _number(additive),
    )


def _norm_line(raw: object) -> NormLine:
    if not isinstance(raw, Mapping):
        raise ValueError("строка металла нормы задана неверно")
    return NormLine(article=str(raw["article"]), kg=_number(raw["kg"]))


def _number(value: object) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError("норма содержит нечисловое поле")
    return float(value)


def _quoted_prices(inputs: PileCostInputs) -> dict[_PriceKey, float]:
    return {(row.mark, row.concrete_grade): _buyer_for(row, inputs) for row in _load_norm_rows()}


def _buyer_for(row: _NormRow, inputs: PileCostInputs) -> float:
    return _cost_for(row, inputs).buyer_price


def _cost_for(row: _NormRow, inputs: PileCostInputs) -> PileCost:
    try:
        recipe = inputs.recipes[row.concrete_grade]
    except KeyError as exc:
        raise ValueError(f"нет рецепта класса {row.concrete_grade}") from exc
    return buyer_price(
        section_cm=row.section_cm,
        volume_m3=row.volume_m3,
        concrete_grade=row.concrete_grade,
        recipe=recipe,
        lines=row.lines,
        prices=inputs.prices,
        overhead=inputs.overhead,
        additive_kg=row.additive_kg,
    )


def _mark_index(rows: tuple[_NormRow, ...]) -> tuple[_MarkNorm, ...]:
    order: list[str] = []
    grades: dict[str, dict[str, _NormRow]] = {}
    section: dict[str, int] = {}
    volume: dict[str, float] = {}
    for row in rows:
        if row.mark not in grades:
            order.append(row.mark)
            grades[row.mark] = {}
            section[row.mark] = row.section_cm
            volume[row.mark] = row.volume_m3
        grades[row.mark][row.concrete_grade] = row
    return tuple(
        _MarkNorm(
            mark=name,
            section_cm=section[name],
            volume_m3=volume[name],
            grades=grades[name],
        )
        for name in order
    )


def _summary(item: _MarkNorm, inputs: PileCostInputs) -> PileCostMarkSummary:
    return PileCostMarkSummary(
        mark=item.mark,
        section_cm=item.section_cm,
        grades=[
            PileCostGradePrice(concrete_grade=grade, buyer_price=_listed_price(item, grade, inputs))
            for grade in GRADE_CODES
        ],
    )


def _listed_price(item: _MarkNorm, grade: str, inputs: PileCostInputs) -> float | None:
    row = item.grades.get(grade)
    if row is None or not _lines_priced(row, inputs):
        return None
    return _buyer_for(row, inputs)


def _lines_priced(row: _NormRow, inputs: PileCostInputs) -> bool:
    known = inputs.prices.prices
    return all(line.article in known for line in row.lines)


def _breakdown(mark: str, marks: tuple[_MarkNorm, ...], inputs: PileCostInputs) -> PileCostMarkBreakdown:
    item = next((row for row in marks if row.mark == mark), None)
    if item is None:
        raise PileCostNotFound()
    return PileCostMarkBreakdown(
        mark=item.mark,
        section_cm=item.section_cm,
        volume_m3=item.volume_m3,
        grades=[_grade_breakdown(grade, _grade_row(item, grade), inputs) for grade in GRADE_CODES],
    )


def _grade_row(item: _MarkNorm, grade: str) -> _NormRow:
    row = item.grades.get(grade)
    if row is None:
        raise ValueError(f"нет класса {grade} у марки {item.mark}")
    return row


def _grade_breakdown(grade: str, row: _NormRow, inputs: PileCostInputs) -> PileCostGradeBreakdown:
    cost = _cost_for(row, inputs)
    return PileCostGradeBreakdown(
        concrete_grade=grade,
        materials_rub=cost.materials_rub,
        expenses_rub=cost.expenses_rub,
        cost=cost.cost,
        buyer_price=cost.buyer_price,
        surcharge_applied=cost.surcharge_applied,
        cement_kg=cost.cement_kg,
        sand_kg=cost.sand_kg,
        crushed_kg=cost.crushed_kg,
        additive_kg=cost.additive_kg,
        cement_rub=cost.cement_rub,
        sand_rub=cost.sand_rub,
        crushed_rub=cost.crushed_rub,
        additive_rub=cost.additive_rub,
        crushed_article=cost.crushed_article,
        lines=[_priced_line(line) for line in cost.lines],
    )


def _priced_line(line: PricedLine) -> PileCostPricedLine:
    return PileCostPricedLine(article=line.article, kg=line.kg, rub=line.rub)


def _inputs_body(inputs: PileCostInputs) -> PileCostInputsBody:
    return PileCostInputsBody(
        prices=[
            PileCostArticlePrice(article=article, price=price)
            for article, price in inputs.prices.prices.items()
        ],
        recipes=[_recipe_body(inputs, grade) for grade in GRADE_CODES],
        overhead=_overhead_body(inputs.overhead),
    )


def _recipe_body(inputs: PileCostInputs, grade: str) -> PileCostRecipeInput:
    recipe = inputs.recipes[grade]
    return PileCostRecipeInput(
        concrete_grade=grade,
        cement_kg_m3=recipe.cement_kg_m3,
        sand_kg_m3=recipe.sand_kg_m3,
        crushed_kg_m3=recipe.crushed_kg_m3,
    )


def _overhead_body(overhead: Overhead) -> PileCostOverheadInput:
    return PileCostOverheadInput(
        general=overhead.general,
        payroll=overhead.payroll,
        forming=overhead.forming,
        other=overhead.other,
    )


def _inputs_from_body(body: PileCostInputsBody) -> PileCostInputs:
    return PileCostInputs(
        prices=PriceBook(_price_map(body)),
        recipes=_recipe_map(body),
        overhead=Overhead(
            general=body.overhead.general,
            payroll=body.overhead.payroll,
            forming=body.overhead.forming,
            other=body.overhead.other,
        ),
    )


def _price_map(body: PileCostInputsBody) -> dict[str, float]:
    prices: dict[str, float] = {}
    for item in body.prices:
        if item.article in prices:
            raise ValueError(f"неизвестный код статьи {item.article}")
        prices[item.article] = item.price
    return prices


def _recipe_map(body: PileCostInputsBody) -> dict[str, Recipe]:
    recipes: dict[str, Recipe] = {}
    for item in body.recipes:
        if item.concrete_grade in recipes:
            raise ValueError(f"класс {item.concrete_grade} не входит в рецепт")
        recipes[item.concrete_grade] = Recipe(
            cement_kg_m3=item.cement_kg_m3,
            sand_kg_m3=item.sand_kg_m3,
            crushed_kg_m3=item.crushed_kg_m3,
        )
    return recipes


def _current_prices(db_path: str) -> dict[_PriceKey, float]:
    conn = sqlite3.connect(db_path)
    try:
        if not _pile_table_exists(conn):
            return {}
        rows = conn.execute("SELECT mark, concrete_grade, price FROM pile_prices").fetchall()
    finally:
        conn.close()
    return {(str(mark), str(grade)): float(price) for mark, grade, price in rows}


def _pile_table_exists(conn: sqlite3.Connection) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'pile_prices'"
    ).fetchone()
    return row is not None


def _preview(
    inputs: PileCostInputs,
    quoted: Mapping[_PriceKey, float],
    current: Mapping[_PriceKey, float],
) -> PileCostPreview:
    diff = diff_price_maps(current, quoted)
    return PileCostPreview(
        inputs_sha256=inputs_sha256(inputs),
        changed=diff.changed,
        new=diff.new,
        missing=diff.missing,
        unchanged=diff.unchanged,
        examples=_examples(diff.examples),
    )


def _examples(examples: PriceDeskDiffExamples) -> PileCostDiffExamples:
    return PileCostDiffExamples(
        changed=[_example(item) for item in examples.changed],
        new=[_example(item) for item in examples.new],
        missing=[_example(item) for item in examples.missing],
        unchanged=[_example(item) for item in examples.unchanged],
    )


def _example(item: DiffExample) -> PileCostDiffExample:
    return PileCostDiffExample(key=item.key, old_price=item.old_price, new_price=item.new_price)


def _write_norm_prices(
    db_path: str,
    quoted: Mapping[_PriceKey, float],
    *,
    written_at: datetime,
) -> None:
    """INSERT OR REPLACE ключей нормы. Строки вне снимка не удаляет."""
    init_pile_prices_schema(db_path)
    list_date = written_at.date().isoformat()
    imported_at = written_at.isoformat(timespec="seconds")
    payload = [
        (mark, grade, price, list_date, imported_at) for (mark, grade), price in quoted.items()
    ]
    conn = _connect(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        conn.executemany(
            """
            INSERT OR REPLACE INTO pile_prices
                (mark, concrete_grade, price, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            payload,
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
