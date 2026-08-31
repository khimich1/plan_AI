from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

from core.march_price_db import (
    get_march_price,
    import_march_prices_from_xlsx,
    normalize_march_mark,
    parse_march_price_rows_from_xlsx,
)


def _write_sample_march_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [1, "Лестничные марши 1ЛМ 27-11-14-4", 13993.72, 14150.79, 14271.10, 14391.41, 14639.53],
        [2, "Лестничные марши 1ЛМ 27-12-14-4", 14430.44, 14592.84, 14717.23, 14841.62, 15098.14],
        [3, "Лестничные марши 1ЛМ 30-11-15-4", 14010.25, 14167.32, 14287.63, 14407.94, 14656.06],
        [4, "Лестничные марши 1ЛМ 30-11-15-4 закладные справа", 14010.25, 14167.32, 14287.63, 14407.94, 14656.06],
        [5, "Лестничные марши 1ЛМ 30-12-15-4  ", 16403.81, 16590.17, 16732.91, 16875.65, 17170.02],
        [6, "Лестничные марши ЛМ 2,8", 15819.88, 16000.91, 16139.57, 16278.24, 16564.20],
        [7, "Лестничные марши ЛМ 2,9", 15819.88, 16000.91, 16139.57, 16278.24, 16564.20],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def test_normalize_march_mark_strips_prefix_and_canon_comma() -> None:
    assert normalize_march_mark("Лестничные марши 1ЛМ 27-11-14-4") == "1ЛМ 27-11-14-4"
    assert normalize_march_mark("ЛМ 2.8") == "ЛМ 2,8"
    assert normalize_march_mark("ЛМ 2,8") == "ЛМ 2,8"
    assert (
        normalize_march_mark("Лестничные марши 1ЛМ 30-11-15-4 закладные справа")
        == "1ЛМ 30-11-15-4 закладные справа"
    )


def test_parse_march_price_rows_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "marches.xlsx"
    _write_sample_march_xlsx(xlsx_path)

    rows = parse_march_price_rows_from_xlsx(str(xlsx_path), preferred_sheet="Прайс")

    assert len(rows) == 35  # 7 × 5
    marks = {r[0] for r in rows}
    assert marks == {
        "1ЛМ 27-11-14-4",
        "1ЛМ 27-12-14-4",
        "1ЛМ 30-11-15-4",
        "1ЛМ 30-11-15-4 закладные справа",
        "1ЛМ 30-12-15-4",
        "ЛМ 2,8",
        "ЛМ 2,9",
    }
    assert ("1ЛМ 27-11-14-4", "B25", 14391.41) in rows
    assert ("ЛМ 2,8", "B25", 16278.24) in rows
    assert all(not m.startswith("Лестничные") for m, _, _ in rows)


def test_import_march_prices_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "marches.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_march_xlsx(xlsx_path)

    inserted = import_march_prices_from_xlsx(
        str(xlsx_path),
        str(db_path),
        preferred_sheet="Прайс",
        price_list_date="2026-08-03",
    )
    assert inserted == 35

    price = get_march_price("1ЛМ 27-11-14-4", "B25", str(db_path))
    assert price == 14391.41

    # Dot form resolves to comma canon
    assert get_march_price("ЛМ 2.8", "B25", str(db_path)) == 16278.24

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM march_prices")
        assert cur.fetchone()[0] == 35
        cur.execute(
            "SELECT price_list_date FROM march_prices WHERE mark = ?",
            ("1ЛМ 27-11-14-4",),
        )
        assert cur.fetchone()[0] == "2026-08-03"
    finally:
        conn.close()
