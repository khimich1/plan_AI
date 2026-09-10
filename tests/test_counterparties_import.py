"""CTR-003: CLI-импорт контрагентов из синтетического xlsx."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from openpyxl import Workbook

from app.repositories.counterparties_repository import CounterpartiesRepository
from scripts.import_counterparties import import_counterparties, render_report
from tests.helpers import kp_db_fixtures as fx

_HEADERS = ["Наименование", "Клиент", "Код", "ИНН", "КПП"]


def _xlsx(tmp_path: Path, rows: list[list], name: str = "cp.xlsx") -> Path:
    path = tmp_path / name
    wb = Workbook()
    ws = wb.active
    ws.append(_HEADERS)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


def test_import_inserts_then_idempotent_rerun(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    xlsx = _xlsx(
        tmp_path,
        [
            ["  СТОЛИЦА ООО  ", "Да", "00-1", "7701000001", "770101001"],
            ["ИП Без ИНН", "Да", "00-2", "", ""],
            ["Поставщик", "Нет", "00-3", "9900000000", "990001001"],
        ],
    )
    first = import_counterparties(xlsx, db)
    assert first.added == 3
    assert first.updated == 0
    assert first.unchanged == 0
    assert first.missing_in_file == 0

    repo = CounterpartiesRepository(db_path=db)
    row = repo.get_by_code_1c("00-1")
    assert row is not None
    assert row["name"] == "СТОЛИЦА ООО"
    assert row["name_normalized"] == "столица ооо"
    assert row["is_client"] == 1
    assert row["source"] == "import"
    empty = repo.get_by_code_1c("00-2")
    assert empty is not None
    assert empty["inn"] is None
    assert empty["kpp"] is None
    supplier = repo.get_by_code_1c("00-3")
    assert supplier is not None
    assert supplier["is_client"] == 0

    second = import_counterparties(xlsx, db)
    assert second.added == 0
    assert second.updated == 0
    assert second.unchanged == 3


def test_delta_file_adds_one_and_does_not_delete(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    file1 = _xlsx(
        tmp_path,
        [["РОМАШКА", "Да", "00-1", "111", ""], ["БЕТА", "Да", "00-2", "222", ""]],
        name="d1.xlsx",
    )
    import_counterparties(file1, db)
    file2 = _xlsx(
        tmp_path,
        [
            ["РОМАШКА", "Да", "00-1", "111", ""],
            ["БЕТА", "Да", "00-2", "222", ""],
            ["СУХОНАИНВЕСТСТРОЙ ООО", "Да", "00-00003431", "3528000000", ""],
        ],
        name="d2.xlsx",
    )
    delta = import_counterparties(file2, db)
    assert delta.added == 1
    assert delta.updated == 0
    repo = CounterpartiesRepository(db_path=db)
    assert repo.get_by_code_1c("00-1") is not None
    assert repo.get_by_code_1c("00-00003431") is not None
    # файл без 00-1 не удаляет запись
    file3 = _xlsx(tmp_path, [["БЕТА", "Да", "00-2", "222", ""]], name="d3.xlsx")
    dropped = import_counterparties(file3, db)
    assert dropped.added == 0
    assert dropped.missing_in_file == 2
    assert repo.get_by_code_1c("00-1") is not None
    assert repo.get_by_code_1c("00-00003431") is not None


def test_dry_run_does_not_write(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    xlsx = _xlsx(tmp_path, [["РОМАШКА", "Да", "00-1", "111", ""]])
    report = import_counterparties(xlsx, db, dry_run=True)
    assert report.added == 1
    with sqlite3.connect(db) as conn:
        assert conn.execute("SELECT COUNT(*) FROM counterparties").fetchone()[0] == 0


def test_inn_conflict_is_reported_and_import_succeeds(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    xlsx = _xlsx(
        tmp_path,
        [
            ["АЛЬФА", "Да", "00-1", "7701000001", ""],
            ["БЕТА", "Да", "00-2", "7701000001", ""],
        ],
    )
    report = import_counterparties(xlsx, db)
    assert report.added == 2
    assert len(report.inn_conflicts) == 1
    inn, codes = report.inn_conflicts[0]
    assert inn == "7701000001"
    assert set(codes) == {"00-1", "00-2"}
    text = render_report(report)
    assert "WARNING" in text
    assert "7701000001" in text


def test_update_on_name_change(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    v1 = _xlsx(tmp_path, [["РОМАШКА", "Да", "00-1", "111", ""]], name="v1.xlsx")
    import_counterparties(v1, db)
    v2 = _xlsx(tmp_path, [["РОМАШКА НОВАЯ", "Да", "00-1", "111", ""]], name="v2.xlsx")
    report = import_counterparties(v2, db)
    assert report.updated == 1
    row = CounterpartiesRepository(db_path=db).get_by_code_1c("00-1")
    assert row is not None
    assert row["name"] == "РОМАШКА НОВАЯ"
