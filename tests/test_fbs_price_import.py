"""FBS-001: FBS price import (sheet Прайс, B7_5/B20/B22_5/B25, skip zeros)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

from core.fbs_price_db import (
    get_fbs_price,
    import_fbs_prices_from_xlsx,
    list_available_grades,
    normalize_fbs_mark_for_lookup,
    parse_fbs_price_rows_from_xlsx,
)


def _write_sample_fbs_xlsx(path: Path) -> None:
    rows = [
        [None, None, None, None, None, None],
        [None, None, None, None, None, None],
        [None, "Наименование", 7.5, 20, 22.5, 25],
        [1, "ФБС 9.3.6-Т", 1640.75, 1731.47, 1759.90, 1788.33],
        [2, "ФБС 12.4.6-Т", 2683.65, 2848.31, 2899.91, 2951.52],
        [3, "ФБС 24.6.6-Т", 7218.26, 7724.67, 7883.37, 0],  # zero B25 skipped
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)
        pd.DataFrame([["noise"]]).to_excel(writer, sheet_name="Цены", index=False, header=False)


def test_normalize_lookup_spaces_and_t() -> None:
    assert normalize_fbs_mark_for_lookup("фбс 9.3.6-т") == "ФБС9.3.6-T"
    assert normalize_fbs_mark_for_lookup("ФБС  9.3.6-Т") == "ФБС9.3.6-T"
    assert normalize_fbs_mark_for_lookup("ФБС9.3.6-T") == "ФБС9.3.6-T"


def test_parse_fbs_price_rows_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "fbs.xlsx"
    _write_sample_fbs_xlsx(xlsx_path)

    rows = parse_fbs_price_rows_from_xlsx(str(xlsx_path), preferred_sheet="Прайс")

    marks = {r[0] for r in rows}
    assert "ФБС 9.3.6-Т" in marks
    assert "ФБС 12.4.6-Т" in marks
    assert "ФБС 24.6.6-Т" in marks

    assert all(price > 0 for _, _, price in rows)
    assert all(g in {"B7_5", "B20", "B22_5", "B25"} for _, g, _ in rows)
    # Zero B25 for 24.6.6 skipped
    assert not any(m == "ФБС 24.6.6-Т" and g == "B25" for m, g, _ in rows)
    # Dense for 9.3.6
    grades_936 = {g for m, g, _ in rows if m == "ФБС 9.3.6-Т"}
    assert grades_936 == {"B7_5", "B20", "B22_5", "B25"}


def test_import_fbs_prices_and_lookup(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "fbs.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_fbs_xlsx(xlsx_path)

    inserted = import_fbs_prices_from_xlsx(
        str(xlsx_path),
        str(db_path),
        preferred_sheet="Прайс",
        price_list_date="2026-08-03",
    )
    assert inserted >= 11

    assert get_fbs_price("ФБС 9.3.6-Т", "B25", str(db_path)) == pytest.approx(1788.33, rel=1e-4) if False else get_fbs_price("ФБС 9.3.6-Т", "B25", str(db_path))
    assert abs(get_fbs_price("ФБС 9.3.6-Т", "B25", str(db_path)) - 1788.33) < 0.01
    assert abs(get_fbs_price("фбс 9.3.6-т", "B7_5", str(db_path)) - 1640.75) < 0.01
    assert get_fbs_price("ФБС 24.6.6-Т", "B25", str(db_path)) is None

    assert list_available_grades("ФБС 9.3.6-Т", str(db_path)) == [
        "B7_5",
        "B20",
        "B22_5",
        "B25",
    ]
    assert list_available_grades("ФБС 24.6.6-Т", str(db_path)) == [
        "B7_5",
        "B20",
        "B22_5",
    ]

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM fbs_prices")
        assert cur.fetchone()[0] == inserted
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='pile_prices'"
        )
        assert cur.fetchone() is None
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='bridge_pile_prices'"
        )
        assert cur.fetchone() is None
    finally:
        conn.close()


def test_real_price_file_if_present() -> None:
    root = Path(__file__).resolve().parents[1]
    xlsx = root / "банк знаний" / "7_5 В прайс на ФБС  от 03.08.2026.xlsx"
    if not xlsx.is_file():
        return
    rows = parse_fbs_price_rows_from_xlsx(str(xlsx), preferred_sheet="Прайс")
    marks = {r[0] for r in rows}
    assert len(marks) >= 14
    assert all(g in {"B7_5", "B20", "B22_5", "B25"} for _, g, _ in rows)
    assert all(p > 0 for _, _, p in rows)
