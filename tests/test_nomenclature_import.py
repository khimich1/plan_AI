"""GPS-016: import-1c выбирает парсер .xls/.xlsx, partial не пишет disappeared."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from openpyxl import Workbook

from app.schemas.nomenclature import Import1cResponse
from app.services.nomenclature_import_service import (
    NomenclatureImportError,
    NomenclatureImportService,
)
from core.nomenclature_guid import ensure_schema, get_by_mark, upsert
from core.nomenclature_sync import sync_pricelist
from core.pricelist_1c_parser import PricelistRow

GUID_A = "11111111-1111-1111-1111-111111111111"
GUID_B = "22222222-2222-2222-2222-222222222222"
GUID_XLSX = "60d7ac63-68b6-11f1-9003-04d9f5387845"


def _seed(path: Path) -> None:
    conn = sqlite3.connect(path)
    try:
        ensure_schema(conn)
        upsert(conn, "pile", "С30.30-3", match_status="missing")
        upsert(
            conn,
            "pile",
            "С70.30-8",
            guid_1c=GUID_B,
            match_status="auto",
        )
        conn.commit()
    finally:
        conn.close()


def _xlsx(
    path: Path,
    *,
    with_filter: bool,
    name: str = "Сваи С 30.30-3",
    guid: str = GUID_XLSX,
    extra_names: list[tuple[str, str]] | None = None,
) -> Path:
    wb = Workbook()
    ws = wb.active
    ws["A2"] = "Параметры:"
    if with_filter:
        ws["A6"] = "Отбор:"
        ws["C6"] = f'Номенклатура Равно "{name}"'
    ws["A8"] = "Номенклатура"
    ws["D8"] = "УИД"
    ws["E8"] = "Код"
    ws["F8"] = "Вес (числитель)"
    ws["H8"] = "Объем (числитель)"
    ws["A9"] = name
    ws["D9"] = guid
    ws["E9"] = "00-00000001"
    ws["F9"] = 1
    ws["H9"] = 0.1
    row = 10
    for extra_name, extra_guid in extra_names or ():
        ws.cell(row, 1, extra_name)
        ws.cell(row, 4, extra_guid)
        ws.cell(row, 5, "00-00000002")
        row += 1
    ws.cell(row, 1, "Итого")
    wb.save(path)
    return path


def test_sync_partial_skips_disappeared(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    _seed(db)
    conn = sqlite3.connect(db)
    try:
        report = sync_pricelist(
            conn,
            [
                PricelistRow(
                    guid=GUID_A,
                    name="Сваи С 30.30-3",
                    price=None,
                    unit="шт",
                    source_file="partial.xlsx",
                    row_index=9,
                )
            ],
            product_kind="pile",
            partial=True,
        )
        leftover = get_by_mark(conn, "pile", "С70.30-8")
    finally:
        conn.close()

    assert report.disappeared == ()
    assert leftover is not None
    assert leftover.guid_1c == GUID_B


def test_xlsx_partial_import_sets_mode(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    _seed(db)
    xlsx = _xlsx(tmp_path / "report.xlsx", with_filter=True)
    service = NomenclatureImportService(db)
    result = service.import_file(xlsx.read_bytes(), "report.xlsx")

    assert isinstance(result, Import1cResponse)
    assert result.mode == "partial"
    assert result.product_kind == "pile"
    assert result.disappeared_count == 0
    assert result.new_guids_count == 1
    conn = sqlite3.connect(db)
    try:
        row = get_by_mark(conn, "pile", "С30.30-3")
        leftover = get_by_mark(conn, "pile", "С70.30-8")
    finally:
        conn.close()
    assert row is not None
    assert row.guid_1c == GUID_XLSX
    assert leftover is not None
    assert leftover.guid_1c == GUID_B


def test_xlsx_full_without_filter_can_mark_disappeared(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    _seed(db)
    xlsx = _xlsx(tmp_path / "full.xlsx", with_filter=False)
    result = NomenclatureImportService(db).import_file(xlsx.read_bytes(), "full.xlsx")
    assert result.mode == "full"
    assert result.disappeared_count >= 1


def test_xlsx_mixed_kinds_require_product_kind(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    _seed(db)
    xlsx = _xlsx(
        tmp_path / "mix.xlsx",
        with_filter=True,
        extra_names=[("Блоки ФБС 9.3.6-Т", GUID_A)],
    )
    try:
        NomenclatureImportService(db).import_file(xlsx.read_bytes(), "mix.xlsx")
        raised = False
    except NomenclatureImportError as exc:
        raised = True
        assert "product_kind" in str(exc)
    assert raised


def test_import_xls_still_returns_full_mode(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    _seed(db)
    service = NomenclatureImportService(db)
    rows = [
        PricelistRow(
            guid=GUID_A,
            name="Сваи С 30.30-3",
            price=None,
            unit="шт",
            source_file="Прайс сваи.xls",
            row_index=3,
        )
    ]
    result = service._sync_and_commit(rows, "pile")
    assert result.mode == "full"
    assert result.new_guids_count == 1


def test_xlsx_import_writes_fbs_weight(tmp_path: Path) -> None:
    db = tmp_path / "pb.db"
    conn = sqlite3.connect(db)
    try:
        ensure_schema(conn)
        upsert(conn, "fbs", "ФБС 9.3.6-Т", match_status="missing")
        conn.commit()
    finally:
        conn.close()
    xlsx = _xlsx(
        tmp_path / "fbs.xlsx",
        with_filter=True,
        name="Блоки ФБС 9.3.6-Т",
        guid=GUID_XLSX,
    )
    # override weight/volume in the helper's defaults (1 / 0.1)
    from openpyxl import load_workbook

    wb = load_workbook(xlsx)
    ws = wb.active
    ws["F9"] = 350
    ws["H9"] = 0.146
    wb.save(xlsx)

    result = NomenclatureImportService(db).import_file(xlsx.read_bytes(), "fbs.xlsx")
    assert result.product_kind == "fbs"
    assert result.weights_updated == 1
    from core.product_weight_catalog import resolve_product_weight_kg

    assert resolve_product_weight_kg("ФБС 9.3.6-Т", "fbs", str(db)) == pytest.approx(350.0)
