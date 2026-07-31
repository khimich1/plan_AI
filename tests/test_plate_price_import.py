from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

from core.price_db import import_from_xlsx, parse_plate_price_rows_from_xlsx


def _write_new_format_xlsx(path: Path) -> None:
    df = pd.DataFrame(
        {
            "Unnamed: 0": ["ПБ 17-12", "ПБ 18-12"],
            "6 нагрузка": [6049, 6389],
            "8 нагрузка": [6049, 6389],
            "10 нагрузка": [6049, 6389],
            "12 нагрузка": [6049, 6389],
        }
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="14.07.2026", index=False, startrow=1)


def test_parse_new_plate_price_xlsx_format(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "plates.xlsx"
    _write_new_format_xlsx(xlsx_path)

    rows = parse_plate_price_rows_from_xlsx(str(xlsx_path), preferred_sheet="14.07.2026")

    assert len(rows) == 8
    assert (17, 8, 6049.0) in rows
    assert (18, 12, 6389.0) in rows


def test_import_new_plate_price_xlsx_replaces_prices(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "plates.xlsx"
    db_path = tmp_path / "pb.db"
    _write_new_format_xlsx(xlsx_path)

    inserted = import_from_xlsx(str(xlsx_path), str(db_path), preferred_sheet="14.07.2026")
    assert inserted == 8

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT price FROM prices WHERE length_dm=? AND load_code=?", (18, 8))
        assert cur.fetchone()[0] == 6389.0
    finally:
        conn.close()
