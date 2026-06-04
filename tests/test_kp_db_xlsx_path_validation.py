"""Tests for safe XLSX path resolution before KP file persistence."""

from __future__ import annotations

import sqlite3

import pytest

from core import kp_db
from core.config.settings import PROJECT_ROOT
from core.kp_file_paths import (
    resolve_allowed_kp_xlsx_path,
    resolve_kp_xlsx_output_path,
    resolve_kp_xlsx_path_for_write,
)
from tests.helpers import kp_db_fixtures as fx


@pytest.fixture
def iso_db(tmp_path):
    return fx.make_iso_db(tmp_path)


def test_resolve_allowed_path_inside_root(tmp_path) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    xlsx = allowed / "offer.xlsx"
    xlsx.write_bytes(b"test-xlsx")

    resolved = resolve_allowed_kp_xlsx_path(
        str(xlsx),
        allowed_roots=[tmp_path],
    )

    assert resolved == xlsx.resolve()


def test_resolve_rejects_path_outside_roots(tmp_path) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    outside = tmp_path / "outside.xlsx"
    outside.write_bytes(b"x")

    assert (
        resolve_allowed_kp_xlsx_path(
            str(outside),
            allowed_roots=[allowed],
        )
        is None
    )


def test_resolve_rejects_dot_dot(tmp_path) -> None:
    assert (
        resolve_allowed_kp_xlsx_path(
            str(tmp_path / ".." / "escape.xlsx"),
            allowed_roots=[tmp_path],
        )
        is None
    )


def test_resolve_output_path_inside_root(tmp_path) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    dest = allowed / "export" / "offer.xlsx"
    resolved = resolve_kp_xlsx_output_path(str(dest), allowed_roots=[tmp_path])
    assert resolved == dest.resolve()


def test_resolve_output_path_rejects_outside_roots(tmp_path) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    outside = tmp_path / "outside.xlsx"
    assert (
        resolve_kp_xlsx_output_path(str(outside), allowed_roots=[allowed]) is None
    )


def test_get_xlsx_file_write_rejects_unsafe_path(iso_db: str, tmp_path) -> None:
    fx.seed_kp_offer(iso_db, 1)
    with sqlite3.connect(iso_db) as conn:
        conn.execute(
            "INSERT INTO kp_files (kp_id, xlsx_file) VALUES (1, ?)",
            (b"blob",),
        )
        conn.commit()
    outside = tmp_path / "outside.xlsx"
    assert kp_db.get_xlsx_file(1, output_path=str(outside), db_path=iso_db) is None
    assert not outside.exists()


def test_get_xlsx_file_write_accepts_safe_path(iso_db: str, tmp_path) -> None:
    from core.config.settings import get_settings

    fx.seed_kp_offer(iso_db, 1)
    with sqlite3.connect(iso_db) as conn:
        conn.execute(
            "INSERT INTO kp_files (kp_id, xlsx_file) VALUES (1, ?)",
            (b"safe-blob",),
        )
        conn.commit()
    drafts = get_settings().drafts_dir
    drafts.mkdir(parents=True, exist_ok=True)
    dest = drafts / f"test_get_xlsx_write_{tmp_path.name}.xlsx"
    try:
        data = kp_db.get_xlsx_file(1, output_path=str(dest), db_path=iso_db)
        assert data == b"safe-blob"
        assert dest.read_bytes() == b"safe-blob"
    finally:
        dest.unlink(missing_ok=True)


def test_save_kp_to_db_stores_blob_for_allowed_path(iso_db: str, tmp_path) -> None:
    from core.config.settings import get_settings

    drafts = get_settings().drafts_dir
    drafts.mkdir(parents=True, exist_ok=True)
    xlsx = drafts / f"test_save_kp_{tmp_path.name}.xlsx"
    xlsx.write_bytes(b"kp-content")

    kp_id = kp_db.save_kp_to_db(
        "01.06.2026",
        [{"name": "ПБ", "qty": 1, "unit_price": 1.0, "length_m": 1.0, "width_m": 1.0}],
        xlsx_file_path=str(xlsx),
        db_path=iso_db,
    )

    with sqlite3.connect(iso_db) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT xlsx_file, file_path FROM kp_files WHERE kp_id = ?",
            (kp_id,),
        )
        row = cur.fetchone()
    assert row is not None
    assert row[0] == b"kp-content"
    assert row[1] == str(xlsx.resolve())
    xlsx.unlink(missing_ok=True)


def test_save_kp_to_db_skips_unsafe_path(iso_db: str, tmp_path) -> None:
    outside = tmp_path / "outside.xlsx"
    outside.write_bytes(b"secret")

    kp_id = kp_db.save_kp_to_db(
        "01.06.2026",
        [{"name": "ПБ", "qty": 1, "unit_price": 1.0, "length_m": 1.0, "width_m": 1.0}],
        xlsx_file_path=str(outside),
        db_path=iso_db,
    )

    with sqlite3.connect(iso_db) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM kp_files WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == 0


def test_save_xlsx_file_rejects_missing_file(iso_db: str) -> None:
    fx.seed_kp_offer(iso_db, 1)
    assert kp_db.save_xlsx_file(1, "/nonexistent/file.xlsx", db_path=iso_db) is False


def test_save_xlsx_file_accepts_file_under_project_root(iso_db: str, tmp_path) -> None:
    fx.seed_kp_offer(iso_db, 1)
    drafts = PROJECT_ROOT / ".app_data" / "drafts"
    drafts.mkdir(parents=True, exist_ok=True)
    xlsx = drafts / f"test_kp_{tmp_path.name}.xlsx"
    xlsx.write_bytes(b"draft-kp")
    try:
        assert kp_db.save_xlsx_file(1, str(xlsx), db_path=iso_db) is True
        with sqlite3.connect(iso_db) as conn:
            cur = conn.cursor()
            cur.execute("SELECT xlsx_file FROM kp_files WHERE kp_id = 1")
            assert cur.fetchone()[0] == b"draft-kp"
    finally:
        if xlsx.is_file():
            xlsx.unlink()
