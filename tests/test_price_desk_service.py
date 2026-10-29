"""PD-007: PriceDeskService preview/apply without HTTP."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd
import pytest

from app.services.price_desk_service import (
    MSG_NEED_SHA,
    MSG_NOT_PILE,
    MSG_SHA_MISMATCH,
    PriceDeskError,
    PriceDeskService,
    file_sha256,
)
from core.pile_price_db import init_pile_prices_schema


def _write_nikita_xlsx(path: Path) -> None:
    rows = [
        [
            None,
            "Наименование",
            "М400",
            "Цены по прайсу, который прислал Никита",
        ],
        [1, "ПБ 17-12-6", 1111, 6049],
        [2, "ПБ 17-12-8", 1111, 6100],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="Прайс", index=False, header=False)


def _write_fbs_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 7.5, 20, 22.5, 25],
        [1, "ФБС 9.3.6-Т", 1640.75, 1731.47, 1759.90, 1788.33],
        [2, "ФБС 12.4.6-Т", 2683.65, 2848.31, 2899.91, 2951.52],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="Прайс", index=False, header=False)


def _write_pile_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [69, "С120.35-12", 11111.11, 22222.22, 33333.33, 44444.44, 55555.55],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="Прайс", index=False, header=False)


def _write_composite_pile_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [1, "Сваи С 60.30-ВС.1", 100.0, 110.0, 120.0, 130.0, 140.0],
        [2, "Сваи С 80.30-НС.1", 200.0, 210.0, 220.0, 230.0, 240.0],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="Прайс", index=False, header=False)


def _count(db: Path, table: str) -> int:
    conn = sqlite3.connect(db)
    try:
        return int(conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
    except sqlite3.OperationalError:
        return 0
    finally:
        conn.close()


def test_preview_does_not_write(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    service = PriceDeskService(db_path=db)
    preview = service.preview(xlsx.read_bytes(), xlsx.name)
    assert preview.product_kind == "plates"
    assert preview.parsed_rows == 2
    assert preview.new == 2
    assert preview.file_sha256 == file_sha256(xlsx.read_bytes())
    assert preview.price_list_date == "2026-08-17"
    assert _count(db, "prices") == 0


def test_apply_writes_then_second_preview_is_unchanged(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    service = PriceDeskService(db_path=db)
    data = xlsx.read_bytes()
    digest = file_sha256(data)
    applied = service.apply(data, xlsx.name, digest)
    assert applied.parsed_rows == 2
    assert _count(db, "prices") == 2
    again = service.preview(data, xlsx.name)
    assert again.changed == 0
    assert again.new == 0
    assert again.unchanged == 2


def test_apply_wrong_sha_does_not_write(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    service = PriceDeskService(db_path=db)
    with pytest.raises(PriceDeskError, match=MSG_SHA_MISMATCH) as exc:
        service.apply(xlsx.read_bytes(), xlsx.name, "ab" * 32)
    assert exc.value.status_code == 409
    assert _count(db, "prices") == 0


def test_apply_without_sha_raises(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    service = PriceDeskService(db_path=db)
    with pytest.raises(PriceDeskError, match=MSG_NEED_SHA):
        service.apply(b"abc", "Прайс ЛМ.xlsx", None)


def test_composite_pile_preview_and_apply(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    xlsx = tmp_path / "Прайс на составные сваи от 07.09.2026.xlsx"
    _write_composite_pile_xlsx(xlsx)
    service = PriceDeskService(db_path=db)
    data = xlsx.read_bytes()
    preview = service.preview(data, xlsx.name)
    assert preview.product_kind == "composite_pile"
    assert preview.parsed_rows == 10
    assert _count(db, "composite_pile_prices") == 0
    applied = service.apply(data, xlsx.name, file_sha256(data))
    assert applied.product_kind == "composite_pile"
    assert _count(db, "composite_pile_prices") == 10


def test_pile_filename_with_fbs_marks_does_not_write_piles(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    xlsx = tmp_path / "Прайс на цельные сваи от 07.09.2026.xlsx"
    _write_fbs_xlsx(xlsx)
    service = PriceDeskService(db_path=db)
    data = xlsx.read_bytes()
    with pytest.raises(PriceDeskError, match=MSG_NOT_PILE) as exc:
        service.preview(data, xlsx.name)
    assert exc.value.status_code == 422
    with pytest.raises(PriceDeskError, match=MSG_NOT_PILE) as apply_exc:
        service.apply(data, xlsx.name, file_sha256(data))
    assert apply_exc.value.status_code == 422
    assert _count(db, "pile_prices") == 0
    assert _count(db, "fbs_prices") == 0


def test_fbs_apply(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    xlsx = tmp_path / "Прайс на ФБС от 07.09.2026.xlsx"
    _write_fbs_xlsx(xlsx)
    service = PriceDeskService(db_path=db)
    data = xlsx.read_bytes()
    result = service.apply(data, xlsx.name, file_sha256(data))
    assert result.product_kind == "fbs"
    assert result.parsed_rows == 8
    assert _count(db, "fbs_prices") == 8


def _seed_pile_prices(db: Path, rows: list[tuple]) -> None:
    init_pile_prices_schema(str(db))
    conn = sqlite3.connect(db)
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


def _pile_rows(db: Path) -> list[tuple]:
    conn = sqlite3.connect(db)
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


def test_pile_preview_parses_sheet_and_apply_keeps_absent_mark(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    kept = ("С30.20-3", "B15", 1000.0, "2026-01-01", "2026-01-01T00:00:00")
    _seed_pile_prices(
        db,
        [
            kept,
            ("С120.35-12", "B15", 1.0, "2026-01-01", "2026-01-01T00:00:00"),
        ],
    )
    before = _pile_rows(db)
    xlsx = tmp_path / "Прайс на цельные сваи от 07.09.2026.xlsx"
    _write_pile_xlsx(xlsx)
    service = PriceDeskService(db_path=db)
    data = xlsx.read_bytes()

    preview = service.preview(data, xlsx.name)
    assert preview.product_kind == "pile"
    assert preview.parsed_rows == 5
    assert preview.price_list_date == "2026-09-07"
    assert preview.new == 4
    assert preview.changed == 1
    assert preview.missing == 1
    assert _pile_rows(db) == before

    applied = service.apply(data, xlsx.name, file_sha256(data))
    assert applied.product_kind == "pile"
    assert applied.parsed_rows == 5
    rows = _pile_rows(db)
    assert len(rows) == 6
    by_key = {
        (mark, grade): (price, list_date, imported_at)
        for mark, grade, price, list_date, imported_at in rows
    }
    assert by_key[("С30.20-3", "B15")] == kept[2:]
    price, list_date, imported_at = by_key[("С120.35-12", "B15")]
    assert price == pytest.approx(11111.11)
    assert list_date == "2026-09-07"
    assert imported_at != "2026-01-01T00:00:00"
    assert by_key[("С120.35-12", "B20")][0] == pytest.approx(22222.22)
    assert by_key[("С120.35-12", "B22_5")][0] == pytest.approx(33333.33)
    assert by_key[("С120.35-12", "B25")][0] == pytest.approx(44444.44)
    assert by_key[("С120.35-12", "B30_granite")][0] == pytest.approx(55555.55)


def test_status_empty_groups(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    status = PriceDeskService(db_path=db).status()
    kinds = [g.product_kind for g in status.groups]
    assert kinds == [
        "plates",
        "fbs",
        "march",
        "step",
        "bridge_pile",
        "pile",
        "composite_pile",
    ]
    assert all(g.row_count == 0 for g in status.groups)
