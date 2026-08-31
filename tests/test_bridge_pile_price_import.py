"""BP-001: bridge pile price import (sheet Прайс, B25/B30, aliases, skip zeros)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

from core.bridge_pile_price_db import (
    get_bridge_pile_price,
    import_bridge_pile_prices_from_xlsx,
    list_available_grades,
    normalize_bridge_pile_mark_for_lookup,
    parse_bridge_pile_price_rows_from_xlsx,
)


def _write_sample_bridge_pile_xlsx(path: Path) -> None:
    rows = [
        [None, None, None, None],
        [None, None, None, None],
        [None, "Наименование", 25, 30],
        [1, "C8-35T1", 35695.27, 0],
        [2, "C8-35T4; C8-35В4", 49813.83, 0],
        [3, "C13-35T4;C13-35В4", 0, 80928.58],
        [4, "C13-40T3", 0, 89879.61],
        [5, "С7-35Т5", 43205.20, 0],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)
        # Ignored sheet — must not be imported
        pd.DataFrame([["noise"]]).to_excel(writer, sheet_name="Цены", index=False, header=False)


def test_normalize_lookup_cyrillic_latin() -> None:
    assert normalize_bridge_pile_mark_for_lookup("с8-35t1") == "C8-35T1"
    assert normalize_bridge_pile_mark_for_lookup("C8-35В4") == "C8-35B4"
    assert normalize_bridge_pile_mark_for_lookup("С7-35Т5") == "C7-35T5"
    assert normalize_bridge_pile_mark_for_lookup("C8-35B4") == "C8-35B4"


def test_parse_bridge_pile_price_rows_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "bridge.xlsx"
    _write_sample_bridge_pile_xlsx(xlsx_path)

    rows = parse_bridge_pile_price_rows_from_xlsx(str(xlsx_path), preferred_sheet="Прайс")

    marks = {r[0] for r in rows}
    assert "C8-35T1" in marks
    assert "C8-35T4" in marks
    assert "C8-35В4" in marks
    assert "C13-35T4" in marks
    assert "C13-35В4" in marks
    assert "С7-35Т5" in marks

    # Zero cells skipped
    assert all(price > 0 for _, _, price, _ in rows)
    assert not any(m == "C8-35T1" and g == "B30" for m, g, _, _ in rows)

    # Alias share variant_group
    t4 = next(r for r in rows if r[0] == "C8-35T4" and r[1] == "B25")
    v4 = next(r for r in rows if r[0] == "C8-35В4" and r[1] == "B25")
    assert t4[2] == v4[2]
    assert t4[3] == v4[3]

    # B30-only row
    assert ("C13-40T3", "B30", 89879.61) == next(
        (m, g, p) for m, g, p, _ in rows if m == "C13-40T3"
    )


def test_import_bridge_pile_prices_and_lookup(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "bridge.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_bridge_pile_xlsx(xlsx_path)

    inserted = import_bridge_pile_prices_from_xlsx(
        str(xlsx_path),
        str(db_path),
        preferred_sheet="Прайс",
        price_list_date="2026-08-03",
    )
    assert inserted >= 7

    assert get_bridge_pile_price("C8-35T1", "B25", str(db_path)) == 35695.27
    # Synonym / cyrillic lookup
    assert get_bridge_pile_price("C8-35В4", "B25", str(db_path)) == 49813.83
    assert get_bridge_pile_price("c8-35t4", "B25", str(db_path)) == 49813.83
    assert get_bridge_pile_price("С7-35Т5", "B25", str(db_path)) == 43205.20
    assert get_bridge_pile_price("C13-40T3", "B30", str(db_path)) == 89879.61
    assert get_bridge_pile_price("C8-35T1", "B30", str(db_path)) is None

    assert list_available_grades("C8-35T1", str(db_path)) == ["B25"]
    assert list_available_grades("C13-40T3", str(db_path)) == ["B30"]
    assert list_available_grades("C8-35В4", str(db_path)) == ["B25"]

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM bridge_pile_prices")
        assert cur.fetchone()[0] == inserted
        # Must not create pile_prices
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='pile_prices'"
        )
        assert cur.fetchone() is None
    finally:
        conn.close()


def test_real_price_file_if_present() -> None:
    root = Path(__file__).resolve().parents[1]
    xlsx = root / "банк знаний" / "Прайс на мостовые сваи от 03.08.2026.xlsx"
    if not xlsx.is_file():
        return
    rows = parse_bridge_pile_price_rows_from_xlsx(str(xlsx), preferred_sheet="Прайс")
    marks = {r[0] for r in rows}
    assert len(marks) >= 64
    assert all(g in {"B25", "B30"} for _, g, _, _ in rows)
    assert all(p > 0 for _, _, p, _ in rows)
