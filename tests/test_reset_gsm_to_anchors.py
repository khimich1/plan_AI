"""Tests for scripts/reset_gsm_to_anchors.py — GSM test-run DB reset."""

from __future__ import annotations

import importlib
import sqlite3
from pathlib import Path

import pytest

from core import kp_db_schema

mod = importlib.import_module("scripts.reset_gsm_to_anchors")


def _fresh_db(tmp_path: Path, name: str = "gsm_reset.db") -> Path:
    db_path = tmp_path / name
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(str(db_path))
    return db_path


def _count(db_path: Path, table: str) -> int:
    with sqlite3.connect(str(db_path)) as conn:
        row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    return int(row[0])


def _seed_two_vehicles_with_anchors(db_path: Path) -> dict[str, int]:
    """Two active vehicles: imported seed + auto waybill + tx/batch + route/card."""
    with sqlite3.connect(str(db_path)) as conn:
        conn.execute(
            """
            INSERT INTO gsm_driver (id, full_name, license_number, is_active)
            VALUES (1, 'Driver One', '111', 1), (2, 'Driver Two', '222', 1)
            """
        )
        conn.execute(
            """
            INSERT INTO gsm_vehicle (
                id, name, plate_number, tank_volume_liters,
                norm_summer, norm_winter, primary_driver_id, is_active
            ) VALUES
                (1, 'Palisade', 'О 521 УХ 44', 70, 14.5, 16.0, 1, 1),
                (2, 'Monjaro', 'О 165 ХУ 44', 60, 9.5, 11.0, 2, 1)
            """
        )
        conn.execute(
            """
            INSERT INTO gsm_fuel_card (id, card_number, vehicle_id, assigned_at)
            VALUES (1, '3005454271', 1, '2026-01-01'),
                   (2, '3005454263', 2, '2026-01-01')
            """
        )
        conn.execute(
            """
            INSERT INTO gsm_route (vehicle_id, addr_a, addr_b, km, frequency)
            VALUES (1, 'A', 'B', 100, 5), (2, 'A', 'C', 50, 3)
            """
        )
        conn.execute(
            """
            INSERT INTO gsm_import_batch (
                id, filename, uploaded_at, rows_total, sum_liters, sum_amount
            ) VALUES
                (1, 'a.xls', '2026-07-01T10:00:00', 1, 40.0, 2000.0),
                (2, 'b.xls', '2026-07-02T10:00:00', 1, 30.0, 1500.0)
            """
        )
        conn.execute(
            """
            INSERT INTO gsm_transaction (
                card_id, ts, service_type, qty_liters, amount,
                raw_address, batch_id
            ) VALUES
                (1, '2026-07-05T12:00:00', 'fuel', 40.0, 2000.0, 'AZS', 1),
                (2, '2026-07-06T12:00:00', 'fuel', 30.0, 1500.0, 'AZS', 2)
            """
        )
        # Imported anchors (keep)
        conn.execute(
            """
            INSERT INTO gsm_waybill (
                id, vehicle_id, date, driver_id, status, source,
                odometer_start, odometer_end, fuel_start, fuel_issued, fuel_end,
                route_json
            ) VALUES
                (10, 1, '2026-06-30', 1, 'confirmed', 'imported',
                 136000, 136331, 8.84, 50.0, 15.84, '[]'),
                (20, 2, '2026-06-30', 2, 'exported', 'imported',
                 61000, 61884, 10.0, 20.0, 7.21, '[]')
            """
        )
        # Auto waybills (delete)
        conn.execute(
            """
            INSERT INTO gsm_waybill (
                id, vehicle_id, date, driver_id, status, source,
                odometer_start, odometer_end, fuel_start, fuel_issued, fuel_end,
                route_json
            ) VALUES
                (11, 1, '2026-07-01', 1, 'exported', 'auto',
                 136331, 136681, 15.84, 40.0, 35.0, '[]'),
                (21, 2, '2026-07-01', 2, 'exported', 'auto',
                 61884, 62000, 7.21, 30.0, 5.0, '[]'),
                (22, 2, '2026-08-01', 2, 'draft', 'auto',
                 62000, 62200, 5.0, 0.0, 2.0, '[]')
            """
        )
        conn.commit()
    return {"v1_anchor": 10, "v2_anchor": 20}


def test_dry_run_does_not_mutate(tmp_path: Path) -> None:
    db = _fresh_db(tmp_path)
    _seed_two_vehicles_with_anchors(db)

    plan = mod.run_reset(db_path=db, apply=False)

    assert len(plan.anchors) == 2
    assert plan.waybills_to_delete == 3
    assert plan.txs_total == 2
    assert plan.batches_total == 2
    assert _count(db, "gsm_waybill") == 5
    assert _count(db, "gsm_transaction") == 2
    assert _count(db, "gsm_import_batch") == 2
    assert not list(tmp_path.glob("*.bak-before-gsm-test-*"))


def test_apply_keeps_imported_anchors_clears_tx(tmp_path: Path) -> None:
    db = _fresh_db(tmp_path)
    ids = _seed_two_vehicles_with_anchors(db)

    plan = mod.run_reset(db_path=db, apply=True)

    assert {a.waybill_id for a in plan.anchors} == {ids["v1_anchor"], ids["v2_anchor"]}
    assert _count(db, "gsm_waybill") == 2
    assert _count(db, "gsm_transaction") == 0
    assert _count(db, "gsm_import_batch") == 0
    assert _count(db, "gsm_route") == 2
    assert _count(db, "gsm_fuel_card") == 2
    assert _count(db, "gsm_vehicle") == 2

    with sqlite3.connect(str(db)) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT id, vehicle_id, date, status, source, odometer_end, fuel_end
            FROM gsm_waybill
            ORDER BY vehicle_id
            """
        ).fetchall()
    assert len(rows) == 2
    assert rows[0]["id"] == ids["v1_anchor"]
    assert rows[0]["status"] == "exported"  # confirmed → exported
    assert rows[0]["source"] == "imported"
    assert rows[0]["date"] == "2026-06-30"
    assert rows[0]["odometer_end"] == 136331
    assert float(rows[0]["fuel_end"]) == pytest.approx(15.84)
    assert rows[1]["id"] == ids["v2_anchor"]
    assert rows[1]["status"] == "exported"
    assert rows[1]["source"] == "imported"

    bak_files = list(tmp_path.glob("*.bak-before-gsm-test-*"))
    assert len(bak_files) == 1
    assert _count(bak_files[0], "gsm_waybill") == 5
    assert _count(bak_files[0], "gsm_transaction") == 2


def test_missing_imported_anchor_aborts_without_changes(tmp_path: Path) -> None:
    db = _fresh_db(tmp_path)
    _seed_two_vehicles_with_anchors(db)
    with sqlite3.connect(str(db)) as conn:
        conn.execute("DELETE FROM gsm_waybill WHERE vehicle_id = 2 AND source = 'imported'")
        conn.commit()

    before_wb = _count(db, "gsm_waybill")
    before_tx = _count(db, "gsm_transaction")

    with pytest.raises(mod.ResetGsmError, match="нет imported-якоря"):
        mod.run_reset(db_path=db, apply=True)

    assert _count(db, "gsm_waybill") == before_wb
    assert _count(db, "gsm_transaction") == before_tx
    assert not list(tmp_path.glob("*.bak-before-gsm-test-*"))


def test_cli_dry_run_exit_zero(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    db = _fresh_db(tmp_path)
    _seed_two_vehicles_with_anchors(db)
    code = mod.main(["--db", str(db)])
    assert code == 0
    out = capsys.readouterr().out
    assert "DRY-RUN" in out
    assert "ничего не записано" in out


def test_cli_apply_exit_zero(tmp_path: Path) -> None:
    db = _fresh_db(tmp_path)
    _seed_two_vehicles_with_anchors(db)
    code = mod.main(["--db", str(db), "--apply"])
    assert code == 0
    assert _count(db, "gsm_waybill") == 2
    assert _count(db, "gsm_transaction") == 0
