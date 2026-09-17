"""PD-002: Nikita column parser for plate price sheet «Прайс»."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from core.plate_nikita_parser import (
    NIKITA_HEADER,
    parse_plate_nikita_rows,
)


def _write_nikita_xlsx(path: Path) -> None:
    rows = [
        [
            None,
            "Наименование",
            "М400",
            "М500",
            "Цены по прайсу, который прислал Никита",
        ],
        [1, "ПБ 17-12-6", 1111, 2222, 6049],
        [2, "ПБ 17-12-8", 1111, 2222, 6100],
        [3, "ПБ 17-12-10", 1111, 2222, 6200],
        [4, "ПБ 18-12-12.5", 1111, 2222, 6389],
        [5, "ПБ 18-12-16", 1111, 2222, 7000],
        [6, "ПБ 18-12-21", 1111, 2222, 8000],
        [7, "ПБ 19-12-6", 1111, 2222, None],
        [8, "ПБ 20-12-8", 1111, 2222, 0],
        [9, "ПБ 21-12-6п", 1111, 2222, 6400],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)
        pd.DataFrame([["noise"]]).to_excel(writer, sheet_name="Цены", index=False, header=False)


def test_parse_nikita_loads_and_skips(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx_path)

    rows = parse_plate_nikita_rows(str(xlsx_path))
    by_key = {(length, load): price for length, load, price in rows}

    assert by_key[(17, 6)] == 6049.0
    assert by_key[(17, 8)] == 6100.0
    assert by_key[(17, 10)] == 6200.0
    assert by_key[(18, 12)] == 6389.0
    assert by_key[(21, 6)] == 6400.0
    assert (18, 16) not in by_key
    assert (18, 21) not in by_key
    assert (19, 6) not in by_key
    assert (20, 8) not in by_key
    # M400/M500 columns must not leak into prices
    assert 1111.0 not in by_key.values()
    assert 2222.0 not in by_key.values()
    assert len(rows) == 5


def test_parse_nikita_empty_without_sheet(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "plates.xlsx"
    pd.DataFrame([["x"]]).to_excel(xlsx_path, sheet_name="Цены", index=False, header=False)
    assert parse_plate_nikita_rows(str(xlsx_path)) == []


def test_parse_nikita_empty_without_column(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "plates.xlsx"
    rows = [
        [None, "Наименование", "6 нагрузка"],
        [1, "ПБ 17-12-6", 6049],
    ]
    pd.DataFrame(rows).to_excel(xlsx_path, sheet_name="Прайс", index=False, header=False)
    assert parse_plate_nikita_rows(str(xlsx_path)) == []


def test_nikita_header_constant() -> None:
    assert NIKITA_HEADER == "цены по прайсу, который прислал никита"


def test_import_nikita_writes_date_and_is_idempotent(tmp_path: Path) -> None:
    import sqlite3

    from core.price_db import import_plate_nikita_from_xlsx, list_plate_prices

    xlsx_path = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    db_path = tmp_path / "pb.db"
    _write_nikita_xlsx(xlsx_path)

    first = import_plate_nikita_from_xlsx(str(xlsx_path), str(db_path))
    assert first == 5
    second = import_plate_nikita_from_xlsx(str(xlsx_path), str(db_path))
    assert second == 5
    assert len(list_plate_prices(str(db_path))) == 5

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT price_list_date, imported_at FROM prices WHERE length_dm=? AND load_code=?",
            (17, 6),
        )
        date_val, imported_at = cur.fetchone()
        assert date_val == "2026-08-17"
        assert imported_at
        cur.execute("SELECT COUNT(*) FROM prices")
        assert cur.fetchone()[0] == 5
    finally:
        conn.close()


def test_import_nikita_keeps_missing_marks(tmp_path: Path) -> None:
    import sqlite3

    from core.price_db import import_plate_nikita_from_xlsx, init_schema

    db_path = tmp_path / "pb.db"
    init_schema(str(db_path))
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            "INSERT INTO prices (length_dm, load_code, price) VALUES (99, 6, 1.0)"
        )
        conn.commit()
    finally:
        conn.close()

    xlsx_path = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx_path)
    import_plate_nikita_from_xlsx(str(xlsx_path), str(db_path))

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT price FROM prices WHERE length_dm=? AND load_code=?", (99, 6))
        assert cur.fetchone()[0] == 1.0
    finally:
        conn.close()
