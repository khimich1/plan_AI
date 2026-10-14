"""CP-110/111: composite pile GUID import from 1C sheet."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.composite_pile_guid_import import (
    import_composite_pile_guids_from_xlsx,
    parse_composite_pile_guid_rows_from_xlsx,
)
from core.composite_pile_kit import import_composite_pile_kit_from_xlsx
from core.nomenclature_guid import PRODUCT_KINDS, ensure_schema, get_by_mark
from core.price_db import _connect

KIT = Path(__file__).parent / "fixtures" / "composite_pile_kit_sample.xlsx"
GUID = Path(__file__).parent / "fixtures" / "composite_pile_guid_sample.xlsx"


@pytest.fixture()
def db_with_kit(tmp_path: Path) -> str:
    db = str(tmp_path / "pb.db")
    import_composite_pile_kit_from_xlsx(str(KIT), db)
    return db


def test_product_kind_includes_composite_pile() -> None:
    assert "composite_pile" in PRODUCT_KINDS


def test_exact_canon_writes_auto(db_with_kit: str) -> None:
    report = import_composite_pile_guids_from_xlsx(str(GUID), db_with_kit)
    assert report.written_auto >= 1
    conn = _connect(db_with_kit)
    try:
        ensure_schema(conn)
        row = get_by_mark(conn, "composite_pile", "С80.30-НС.1")
        assert row is not None
        assert row.match_status == "auto"
        assert row.guid_1c == "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
        assert row.guid_1c_u is None
    finally:
        conn.close()


def test_duplicate_guid_is_ambiguous(db_with_kit: str) -> None:
    import_composite_pile_guids_from_xlsx(str(GUID), db_with_kit)
    conn = _connect(db_with_kit)
    try:
        ensure_schema(conn)
        row = get_by_mark(conn, "composite_pile", "С60.30-ВС.1")
        assert row is not None
        assert row.match_status == "ambiguous"
        assert row.guid_1c is None
    finally:
        conn.close()


def test_non_series_names_not_written(db_with_kit: str) -> None:
    import_composite_pile_guids_from_xlsx(str(GUID), db_with_kit)
    conn = _connect(db_with_kit)
    try:
        ensure_schema(conn)
        assert get_by_mark(conn, "composite_pile", "С140.30-С") is None
        # комплект / garbage never become marks
        cur = conn.execute(
            "SELECT mark FROM nomenclature_guid WHERE product_kind = 'composite_pile'"
        )
        marks = {r[0] for r in cur.fetchall()}
        assert all("Комплект" not in m for m in marks)
        assert "Свая гранитная особая" not in marks
    finally:
        conn.close()


def test_refuses_pricelist_xls_source(tmp_path: Path, db_with_kit: str) -> None:
    bad = tmp_path / "Прайс сваи составные.xls"
    bad.write_bytes(b"fake")
    with pytest.raises(ValueError, match="не используется"):
        parse_composite_pile_guid_rows_from_xlsx(str(bad))
