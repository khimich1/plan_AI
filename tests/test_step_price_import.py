from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

from core.step_price_db import (
    extract_step_mark,
    get_step_price,
    import_step_prices_from_xlsx,
    normalize_step_mark,
    parse_step_price_rows_from_xlsx,
)

# Path to real price list (optional integration check)
_REAL_XLSX = (
    Path(__file__).resolve().parents[1]
    / "банк знаний"
    / "Прайс на лестничные ступени от 03.08.2026.xlsx"
)


def _write_sample_step_xlsx(path: Path) -> None:
    rows = [
        [None, None, None],
        [None, None, None],
        [None, "Наименование", 15],
        [1, "Лестничные ступени ЛС11", 1409.908359678],
        [2, "Лестничные ступени ЛС14-1лев", 1815.586530576],
        [3, "Лестничные ступени  ЛС14-Б", 1564.62029178],  # extra space before mark
        [4, "Лестничные ступени ЛС11-Б-1", 1420.991597298],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def test_extract_step_mark_from_full_name() -> None:
    assert extract_step_mark("Лестничные ступени ЛС11") == "ЛС11"
    assert extract_step_mark("Лестничные ступени  ЛС14-Б") == "ЛС14-Б"
    assert extract_step_mark("лс12-2лев") == "ЛС12-2ЛЕВ"
    assert normalize_step_mark("ЛС 14-1лев") == "ЛС14-1ЛЕВ"


def test_parse_step_price_rows_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "steps.xlsx"
    _write_sample_step_xlsx(xlsx_path)

    rows = parse_step_price_rows_from_xlsx(str(xlsx_path), preferred_sheet="Прайс")

    marks = {r[0] for r in rows}
    assert marks == {"ЛС11", "ЛС14-1ЛЕВ", "ЛС14-Б", "ЛС11-Б-1"}
    by_mark = {r[0]: r[1] for r in rows}
    assert by_mark["ЛС11"] == 1409.908359678
    assert by_mark["ЛС14-Б"] == 1564.62029178


def test_import_step_prices_from_xlsx(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "steps.xlsx"
    db_path = tmp_path / "pb.db"
    _write_sample_step_xlsx(xlsx_path)

    inserted = import_step_prices_from_xlsx(
        str(xlsx_path),
        str(db_path),
        preferred_sheet="Прайс",
        price_list_date="2026-08-03",
    )
    assert inserted == 4

    price = get_step_price("ЛС11", str(db_path))
    assert price == 1409.908359678
    assert get_step_price("лс14-1лев", str(db_path)) == 1815.586530576

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT price_list_date, display_name FROM step_prices WHERE mark = ?", ("ЛС11",))
        row = cur.fetchone()
        assert row[0] == "2026-08-03"
        assert "ЛС11" in row[1]
    finally:
        conn.close()


def test_real_price_list_has_at_least_42_marks() -> None:
    if not _REAL_XLSX.is_file():
        return  # skip silently if knowledge-bank file absent in CI
    rows = parse_step_price_rows_from_xlsx(str(_REAL_XLSX), preferred_sheet="Прайс")
    marks = {r[0] for r in rows}
    assert len(marks) >= 42
    assert "ЛС11" in marks
    assert "ЛС14-1ЛЕВ" in marks or "ЛС14-1лев" in {m for m in marks}
    # normalize keeps upper; real extract uses normalize_step_mark → ЛЕВ upper
    assert any(m.startswith("ЛС22") for m in marks)
