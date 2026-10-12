"""PC-004: превью и запись цен цельных свай. Превью не пишет pile_prices."""

from __future__ import annotations

import json
import sqlite3
from datetime import date
from pathlib import Path

import pytest

from app.schemas.pile_cost import PileCostApplyRequest
from app.services.pile_cost_service import (
    MSG_NEED_SHA,
    MSG_STALE,
    PileCostConflict,
    PileCostService,
    inputs_sha256,
)
from core.pile_cost import (
    ARTICLE_CEMENT,
    NormLine,
    PriceBook,
    buyer_price,
)
from core.pile_cost_inputs import PileCostInputs, load_pile_cost_inputs, save_pile_cost_inputs
from core.pile_price_db import GRADE_CODES, init_pile_prices_schema

_NORM_PATH = Path(__file__).resolve().parents[1] / "core" / "pile_cost_norm.json"
_NORM_ROWS = json.loads(_NORM_PATH.read_text(encoding="utf-8"))["rows"]
_EXCLUDED_MARK = "С130.35-13А**"
_S30 = "С30.20-3"
_S30_BUYER = {
    "B15": 3513.2818728,
    "B20": 3538.3051464,
    "B22_5": 3559.5848256,
    "B25": 3580.8645048,
    "B30_granite": 3655.4166552,
}
_TOLERANCE = 0.01
_INPUT_TABLES = (
    "pile_cost_article_price",
    "pile_cost_recipe",
    "pile_cost_overhead",
)


def _db(tmp_path: Path) -> str:
    return str(tmp_path / "pb.db")


def _norm_row(mark: str, grade: str) -> dict[str, object]:
    found = [row for row in _NORM_ROWS if row["mark"] == mark and row["concrete_grade"] == grade]
    assert len(found) == 1
    return found[0]


def _formula_price(row: dict[str, object], inputs: PileCostInputs) -> float:
    grade = str(row["concrete_grade"])
    lines = tuple(NormLine(str(line["article"]), float(line["kg"])) for line in row["lines"])  # type: ignore[index]
    additive = row.get("additive_kg")
    return buyer_price(
        section_cm=int(row["section_cm"]),  # type: ignore[arg-type]
        volume_m3=float(row["volume_m3"]),  # type: ignore[arg-type]
        concrete_grade=grade,
        recipe=inputs.recipes[grade],
        lines=lines,
        prices=inputs.prices,
        overhead=inputs.overhead,
        additive_kg=None if additive is None else float(additive),  # type: ignore[arg-type]
    ).buyer_price


def _articles_in_norm() -> set[str]:
    found: set[str] = set()
    for row in _NORM_ROWS:
        for line in row["lines"]:
            found.add(str(line["article"]))
    return found


def _with_norm_articles(db_path: str) -> PileCostInputs:
    """Дописывает недостающие статьи снимка. Цены 06.10.2026 не затирает."""
    loaded = load_pile_cost_inputs(db_path)
    prices = dict(loaded.prices.prices)
    for article in sorted(_articles_in_norm()):
        prices.setdefault(article, 1.0)
    stored = PileCostInputs(prices=PriceBook(prices), recipes=loaded.recipes, overhead=loaded.overhead)
    save_pile_cost_inputs(stored, db_path)
    return load_pile_cost_inputs(db_path)


def _table_exists(db_path: str, name: str) -> bool:
    conn = sqlite3.connect(db_path)
    try:
        row = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (name,),
        ).fetchone()
        return row is not None
    finally:
        conn.close()


def _pile_rows(db_path: str) -> list[tuple] | None:
    if not _table_exists(db_path, "pile_prices"):
        return None
    conn = sqlite3.connect(db_path)
    try:
        return conn.execute(
            """
            SELECT mark, concrete_grade, price, price_list_date, imported_at
            FROM pile_prices
            ORDER BY mark, concrete_grade
            """
        ).fetchall()
    finally:
        conn.close()


def _seed_pile_prices(db_path: str, rows: list[tuple]) -> None:
    init_pile_prices_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.executemany(
            """
            INSERT INTO pile_prices
                (mark, concrete_grade, price, price_list_date, imported_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            rows,
        )
        conn.commit()
    finally:
        conn.close()


def _input_rows(db_path: str) -> dict[str, list[tuple]]:
    conn = sqlite3.connect(db_path)
    try:
        dumped: dict[str, list[tuple]] = {}
        for table in _INPUT_TABLES:
            dumped[table] = conn.execute(f"SELECT * FROM {table} ORDER BY 1").fetchall()
        return dumped
    finally:
        conn.close()


def _columns(db_path: str, table: str) -> list[str]:
    conn = sqlite3.connect(db_path)
    try:
        return [str(row[1]) for row in conn.execute(f"PRAGMA table_info({table})")]
    finally:
        conn.close()


def test_preview_does_not_change_pile_prices(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    inputs = _with_norm_articles(db_path)
    expected = _formula_price(_norm_row(_S30, "B15"), inputs)
    _seed_pile_prices(
        db_path,
        [(_S30, "B15", expected, "2026-10-06", "2026-10-06T12:00:00")],
    )
    before_prices = _pile_rows(db_path)
    before_inputs = _input_rows(db_path)
    before_norm = _NORM_PATH.read_bytes()

    preview = PileCostService(db_path).preview()

    assert preview.inputs_sha256 == inputs_sha256(inputs)
    assert preview.unchanged >= 1
    assert _pile_rows(db_path) == before_prices
    assert _input_rows(db_path) == before_inputs
    assert _NORM_PATH.read_bytes() == before_norm


def test_preview_without_article_price_writes_nothing(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    load_pile_cost_inputs(db_path)
    _seed_pile_prices(
        db_path,
        [(_EXCLUDED_MARK, "B15", 60980.85, "2026-10-06", "2026-10-06T12:00:00")],
    )
    before = _pile_rows(db_path)

    with pytest.raises(ValueError, match="нет цены статьи"):
        PileCostService(db_path).preview()

    assert _pile_rows(db_path) == before
    assert "volume_m3" not in _columns(db_path, "pile_cost_recipe")
    assert "kg" not in _columns(db_path, "pile_cost_article_price")


def test_apply_without_article_price_writes_nothing(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    inputs = load_pile_cost_inputs(db_path)
    _seed_pile_prices(
        db_path,
        [(_EXCLUDED_MARK, "B15", 60980.85, "2026-10-06", "2026-10-06T12:00:00")],
    )
    before = _pile_rows(db_path)

    with pytest.raises(ValueError, match="нет цены статьи"):
        PileCostService(db_path).apply(inputs_sha256(inputs))

    assert _pile_rows(db_path) == before


def test_foreign_sha256_does_not_write(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    _with_norm_articles(db_path)
    _seed_pile_prices(
        db_path,
        [(_S30, "B15", 1.0, "2026-10-06", "2026-10-06T12:00:00")],
    )
    before = _pile_rows(db_path)
    service = PileCostService(db_path)

    with pytest.raises(PileCostConflict, match=MSG_STALE) as caught:
        service.apply("ab" * 32)

    assert caught.value.status_code == 409
    assert _pile_rows(db_path) == before


def test_apply_without_sha256_does_not_write(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    _with_norm_articles(db_path)
    _seed_pile_prices(
        db_path,
        [(_EXCLUDED_MARK, "B15", 100.0, "2026-10-06", "2026-10-06T12:00:00")],
    )
    before = _pile_rows(db_path)

    with pytest.raises(PileCostConflict, match=MSG_NEED_SHA) as caught:
        PileCostService(db_path).apply("  ")

    assert caught.value.status_code == 409
    assert _pile_rows(db_path) == before


def test_stale_sha_after_cement_change_does_not_write(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    inputs = _with_norm_articles(db_path)
    service = PileCostService(db_path)
    preview = service.preview()
    prices = dict(inputs.prices.prices)
    prices[ARTICLE_CEMENT] = 8.25
    save_pile_cost_inputs(
        PileCostInputs(prices=PriceBook(prices), recipes=inputs.recipes, overhead=inputs.overhead),
        db_path,
    )
    before = _pile_rows(db_path)

    with pytest.raises(PileCostConflict, match=MSG_STALE):
        service.apply(preview.inputs_sha256)

    assert _pile_rows(db_path) == before
    assert service.preview().inputs_sha256 != preview.inputs_sha256


def test_apply_writes_s30_prices_and_keeps_mark_outside_norm(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    inputs = _with_norm_articles(db_path)
    assert inputs.prices.unit_price(ARTICLE_CEMENT) == 7.54
    expected = {
        grade: _formula_price(_norm_row(_S30, grade), inputs) for grade in GRADE_CODES
    }
    for grade, price in expected.items():
        assert price == pytest.approx(_S30_BUYER[grade], abs=_TOLERANCE)

    excluded = [
        (_EXCLUDED_MARK, "B15", 60980.85, "2026-10-01", "2026-10-01T08:00:00"),
        (_EXCLUDED_MARK, "B20", 61000.0, "2026-10-01", "2026-10-01T08:00:00"),
    ]
    _seed_pile_prices(
        db_path,
        excluded + [(_S30, "B15", 1.0, "2026-10-06", "2026-10-06T12:00:00")],
    )
    before_inputs = _input_rows(db_path)
    before_norm = _NORM_PATH.read_bytes()
    service = PileCostService(db_path)
    digest = service.preview().inputs_sha256

    applied = service.apply(PileCostApplyRequest(inputs_sha256=digest.upper()).inputs_sha256)

    assert applied.inputs_sha256 == digest
    assert applied.missing == 2
    assert _input_rows(db_path) == before_inputs
    assert _NORM_PATH.read_bytes() == before_norm

    rows = _pile_rows(db_path)
    assert rows is not None
    by_key = {(mark, grade): (price, list_date, imported_at) for mark, grade, price, list_date, imported_at in rows}
    for grade in GRADE_CODES:
        price, list_date, imported_at = by_key[(_S30, grade)]
        assert price == expected[grade]
        assert price != round(price)
        assert list_date == date.today().isoformat()
        assert imported_at is not None
    assert by_key[(_EXCLUDED_MARK, "B15")] == (60980.85, "2026-10-01", "2026-10-01T08:00:00")
    assert by_key[(_EXCLUDED_MARK, "B20")] == (61000.0, "2026-10-01", "2026-10-01T08:00:00")
    assert len(rows) == len(_NORM_ROWS) + 2

    again = service.preview()
    assert again.changed == 0
    assert again.new == 0
    assert again.missing == 2
    assert again.unchanged == len(_NORM_ROWS)
    assert "volume_m3" not in _columns(db_path, "pile_prices")
    assert "kg" not in _columns(db_path, "pile_prices")


def test_diff_threshold_and_example_limit(tmp_path: Path) -> None:
    db_path = _db(tmp_path)
    inputs = _with_norm_articles(db_path)
    b15 = _formula_price(_norm_row(_S30, "B15"), inputs)
    b20 = _formula_price(_norm_row(_S30, "B20"), inputs)
    b22 = _formula_price(_norm_row(_S30, "B22_5"), inputs)
    _seed_pile_prices(
        db_path,
        [
            (_S30, "B15", b15 + 0.004, "2026-10-06", "2026-10-06T12:00:00"),
            (_S30, "B20", b20 + 0.006, "2026-10-06", "2026-10-06T12:00:00"),
            (_S30, "B22_5", b22, "2026-10-06", "2026-10-06T12:00:00"),
            (_EXCLUDED_MARK, "B15", 100.0, "2026-10-06", "2026-10-06T12:00:00"),
        ],
    )

    preview = PileCostService(db_path).preview()

    assert preview.changed == 1
    assert preview.unchanged == 2
    assert preview.missing == 1
    assert preview.new == len(_NORM_ROWS) - 3
    assert len(preview.examples.changed) == 1
    assert len(preview.examples.unchanged) == 2
    assert len(preview.examples.missing) == 1
    assert len(preview.examples.new) == 30
    assert preview.examples.changed[0].key == f"{_S30} / B20"
    assert preview.examples.changed[0].old_price == pytest.approx(b20 + 0.006)
    assert preview.examples.changed[0].new_price == b20
    assert {item.key for item in preview.examples.unchanged} == {
        f"{_S30} / B15",
        f"{_S30} / B22_5",
    }
    assert preview.examples.missing[0].key == f"{_EXCLUDED_MARK} / B15"
    assert preview.examples.missing[0].old_price == 100.0
    assert preview.examples.missing[0].new_price is None
