from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

from core.pile_price_db import (
    get_pile_price,
    import_pile_prices_from_xlsx,
    parse_pile_price_rows_from_xlsx,
)


def _write_sample_pile_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [None, "35 СЕЧЕНИЕ", None, None, None, None, None],
        [69, "С120.35-12", 43760.30608943999, 44108.14723836479, 44371.0862132184, 44634.02518807199, 46159.3673880192],
        [91, "С120.35-13и", 67512.26545487999, 67860.1066038048, 68123.04557865839, 68385.984553512, 69911.32675345919],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def test_parse_pile_price_rows_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "piles.xlsx"
    _write_sample_pile_xlsx(xlsx_path)

    rows = parse_pile_price_rows_from_xlsx(str(xlsx_path), preferred_sheet="Прайс")

    assert len(rows) == 10
    assert ("С120.35-12", "B25", 44634.02518807199) in rows
    assert ("С120.35-13и", "B30_granite", 69911.32675345919) in rows


def test_import_pile_prices_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "piles.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_pile_xlsx(xlsx_path)

    inserted = import_pile_prices_from_xlsx(
        str(xlsx_path),
        str(db_path),
        preferred_sheet="Прайс",
        price_list_date="2026-07-27",
    )
    assert inserted == 10

    price = get_pile_price("С120.35-12", "B25", str(db_path))
    assert price == 44634.02518807199

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT price_list_date FROM pile_prices WHERE mark = ?", ("С120.35-12",))
        assert cur.fetchone()[0] == "2026-07-27"
    finally:
        conn.close()
