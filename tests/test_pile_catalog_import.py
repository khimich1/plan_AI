"""SHIP-100: парсинг прайса свай + upsert каталога (идемпотентность)."""

from __future__ import annotations

import sqlite3

import pytest
from openpyxl import Workbook

from core.pile_catalog import (
    PileCatalogEntry,
    parse_pile_catalog_from_xlsx,
    parse_pile_mark,
    upsert_pile_catalog,
)
from tests.helpers import kp_db_fixtures as fx

QUIRK_MARK = "С137,5.40"


def _build_price_xlsx(tmp_path, rows: list[tuple]) -> str:
    """Лист «Вес и объем»: заголовок в строке 2, данные со строки 3."""
    path = str(tmp_path / "price.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Вес и объем"
    ws.append(["Прайс цельных свай"])
    ws.append(["Марка", "Объем, м3", "Вес, кг", "Шт. на а/м 20тн"])
    for row in rows:
        ws.append(list(row))
    wb.save(path)
    return path


def _standard_rows(count: int = 43) -> list[tuple]:
    rows = [(f"С{30 + i}.30", 0.1, 1000 + i, 10) for i in range(count)]
    rows.append((QUIRK_MARK, 0.55, 2750, "-"))
    return rows


@pytest.mark.parametrize(
    ("mark", "expected_length", "expected_section"),
    [
        ("С60.30", 6.0, 300),
        (QUIRK_MARK, 13.75, 400),
        ("С120.35", 12.0, 350),
        ("мусор", None, None),
        ("С60", None, None),
    ],
)
def test_parse_pile_mark(mark, expected_length, expected_section) -> None:
    assert parse_pile_mark(mark) == (expected_length, expected_section)


def test_parse_full_sheet_44_rows_and_quirks(tmp_path) -> None:
    xlsx = _build_price_xlsx(
        tmp_path,
        _standard_rows()
        + [
            (QUIRK_MARK, 0.55, 2750, "-"),  # дубликат марки — пропускается
            ("С99.30", None, None, 5),  # без веса — пропускается
            (None, 0.1, 100, 5),  # пустая марка — пропускается
        ],
    )
    entries = parse_pile_catalog_from_xlsx(xlsx)
    assert len(entries) == 44
    quirk = next(entry for entry in entries if entry.mark == QUIRK_MARK)
    assert quirk.length_m == pytest.approx(13.75)
    assert quirk.section_mm == 400
    assert quirk.weight_kg == pytest.approx(2750.0)
    assert quirk.pcs_per_20t is None  # "-" → NULL


def test_parse_missing_sheet_raises(tmp_path) -> None:
    xlsx = _build_price_xlsx(tmp_path, _standard_rows(1))
    with pytest.raises(ValueError, match="не найден"):
        parse_pile_catalog_from_xlsx(xlsx, sheet="Нет такого")


def test_upsert_idempotent(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    entries = [
        PileCatalogEntry(
            mark="С60.30",
            length_m=6.0,
            section_mm=300,
            volume_m3=0.216,
            weight_kg=1060.0,
            pcs_per_20t=18,
        ),
        PileCatalogEntry(
            mark=QUIRK_MARK,
            length_m=13.75,
            section_mm=400,
            volume_m3=0.55,
            weight_kg=2750.0,
            pcs_per_20t=None,
        ),
    ]
    inserted, updated = upsert_pile_catalog(db_path, entries)
    assert (inserted, updated) == (2, 0)
    inserted, updated = upsert_pile_catalog(db_path, entries)
    assert (inserted, updated) == (0, 2)

    changed = [
        PileCatalogEntry(
            mark="С60.30",
            length_m=6.0,
            section_mm=300,
            volume_m3=0.216,
            weight_kg=1111.0,
            pcs_per_20t=17,
        )
    ]
    inserted, updated = upsert_pile_catalog(db_path, changed)
    assert (inserted, updated) == (0, 1)
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT weight_kg, pcs_per_20t FROM pile_catalog WHERE mark = ?",
            ("С60.30",),
        ).fetchone()
        total = conn.execute("SELECT COUNT(*) FROM pile_catalog").fetchone()[0]
    assert row == (1111.0, 17)
    assert total == 2


def test_end_to_end_parse_then_upsert(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    xlsx = _build_price_xlsx(tmp_path, _standard_rows())
    entries = parse_pile_catalog_from_xlsx(xlsx)
    inserted, updated = upsert_pile_catalog(db_path, entries)
    assert (inserted, updated) == (44, 0)
    inserted, updated = upsert_pile_catalog(db_path, entries)
    assert (inserted, updated) == (0, 44)
    with sqlite3.connect(db_path) as conn:
        total = conn.execute("SELECT COUNT(*) FROM pile_catalog").fetchone()[0]
    assert total == 44
