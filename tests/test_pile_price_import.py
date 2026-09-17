from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd

from core.pile_price_db import (
    get_pile_price,
    import_pile_prices_from_xlsx,
    init_pile_prices_schema,
    parse_pile_price_rows_from_xlsx,
)


def _seed_pile_prices(db_path: Path, rows: list[tuple[str, str, float]]) -> None:
    init_pile_prices_schema(str(db_path))
    conn = sqlite3.connect(db_path)
    try:
        conn.executemany(
            "INSERT INTO pile_prices (mark, concrete_grade, price) VALUES (?, ?, ?)",
            rows,
        )
        conn.commit()
    finally:
        conn.close()


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


def test_get_pile_price_matches_spaced_and_latin_mark(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _seed_pile_prices(db_path, [("С 140.30-10", "B25", 38976.66)])

    assert get_pile_price("C140.30-10", "B25", str(db_path)) == 38976.66
    assert get_pile_price("С140.30-10", "B25", str(db_path)) == 38976.66
    assert get_pile_price("С 140.30-10", "B25", str(db_path)) == 38976.66


def test_get_pile_price_matches_latin_c_to_cyrillic(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _seed_pile_prices(db_path, [("С120.35-12", "B25", 44634.03)])

    assert get_pile_price("C120.35-12", "B25", str(db_path)) == 44634.03


def test_get_pile_price_exact_type_a_does_not_use_compact_duplicate(
    tmp_path: Path,
) -> None:
    db_path = tmp_path / "pb.db"
    spaced = 'Тип "А" С120.35-13'
    compact = 'Тип"А"С120.35-13'
    _seed_pile_prices(
        db_path,
        [
            (spaced, "B25", 70448.42),
            (compact, "B25", 79910.61),
        ],
    )

    assert get_pile_price(spaced, "B25", str(db_path)) == 70448.42
    assert get_pile_price(compact, "B25", str(db_path)) == 79910.61
    assert get_pile_price('Тип"А" С120.35-13', "B25", str(db_path)) is None


def test_get_pile_price_unknown_mark_returns_none(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _seed_pile_prices(db_path, [("С120.35-12", "B25", 44634.03)])

    assert get_pile_price("С120.35-99", "B25", str(db_path)) is None


def test_get_pile_price_u_suffix_and_fractional_load(tmp_path: Path) -> None:
    db_path = tmp_path / "pb.db"
    _seed_pile_prices(
        db_path,
        [
            ("С110.30-9", "B25", 10300.0),
            ("С110.40-8", "B25", 11300.0),
            ("С120.35-13и", "B25", 68385.98),
        ],
    )
    assert get_pile_price("С110.30-9у", "B25", str(db_path)) == 10300.0
    assert get_pile_price("C 110.30-9.1у", "B25", str(db_path)) == 10300.0
    assert get_pile_price("C110.40-8.1", "B25", str(db_path)) == 11300.0
    assert get_pile_price("С120.35-13и", "B25", str(db_path)) == 68385.98
    assert get_pile_price("С120.35-13", "B25", str(db_path)) is None
    assert get_pile_price("С110.30-6у", "B25", str(db_path)) is None
