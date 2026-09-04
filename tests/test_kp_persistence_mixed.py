"""MNA-302 / MNA-303 / MNA-304: mixed KP persistence create + read + update sync.

MNA-302 (create):
- Plates→Piles order → rows in both ``kp_plates`` and ``kp_piles``
- ``position_number`` 1..N chronological across tables
- ``kp_meta.product_type == "mixed"`` when ≥2 types; single type unchanged
- Persist ``line_id`` on line rows

MNA-303 (read):
- ``get_kp_by_id`` loads all type tables for mixed; stamps ``product_type`` on each line
- Chronological order by ``position_number`` (typed arrays + ``order_data_from_kp_info``)

MNA-304 (update; archive-edit override):
- ``update_kp_from_order_data`` keeps same ``kp_id``; sync by ``line_id``
- Status gate: only ``в архиве``
- Matching ``line_id`` preserves ``kp_plates.id`` (production-safe)
"""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.kp.offers_read import get_kp_by_id
from core.kp_db_schema import init_schema
from core.kp_order_data import order_data_from_kp_info
from core.kp_persistence_service import KpPersistenceService


@pytest.fixture()
def db_path(tmp_path: Path) -> str:
    path = str(tmp_path / "plita.db")
    init_schema(path)
    return path


def _plate_line(**overrides: object) -> dict:
    base = {
        "line_id": "ln_plate_1",
        "product_type": "plates",
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
        "length_dm_raw": "60",
        # Explicit grade so resolve_concrete_grade short-circuits (no pb.db).
        "concrete_grade": "М500",
    }
    base.update(overrides)
    return base


def _pile_line(**overrides: object) -> dict:
    base = {
        "line_id": "ln_pile_1",
        "product_type": "piles",
        "product_kind": "pile",
        "name": "С120.35-12",
        "mark": "С120.35-12",
        "concrete_grade": "B25",
        "qty": 2,
        "unit_price": 44634.03,
    }
    base.update(overrides)
    return base


def _meta_product_type(db_path: str, kp_id: int) -> str:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = ?", (kp_id,))
        row = cur.fetchone()
        assert row is not None
        return str(row[0])


def _rows(db_path: str, table: str, kp_id: int) -> list[sqlite3.Row]:
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(
            f"SELECT * FROM {table} WHERE kp_id = ? ORDER BY position_number, id",
            (kp_id,),
        )
        return list(cur.fetchall())


# --- Multi-table create -------------------------------------------------------


def test_save_mixed_plates_then_piles_writes_both_tables(db_path: str) -> None:
    """Plates→Piles order_data must land in kp_plates and kp_piles (one KP)."""
    order = [
        _plate_line(line_id="ln_a", name="ПБ 60-12-8п"),
        _pile_line(line_id="ln_b", mark="С120.35-12"),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Mixed Client",
        db_path=db_path,
    )

    plates = _rows(db_path, "kp_plates", kp_id)
    piles = _rows(db_path, "kp_piles", kp_id)
    assert len(plates) == 1
    assert len(piles) == 1
    assert plates[0]["plate_name"] == "ПБ 60-12-8п"
    assert piles[0]["mark"] == "С120.35-12"
    assert plates[0]["qty"] == 1
    assert piles[0]["qty"] == 2


def test_save_mixed_assigns_chronological_position_numbers(db_path: str) -> None:
    """position_number is 1..N across tables in order_data order (not per-table)."""
    order = [
        _plate_line(line_id="ln_1", name="ПБ 60-12-8п"),
        _pile_line(line_id="ln_2", mark="С120.35-12"),
        _plate_line(
            line_id="ln_3",
            name="ПБ 72-12-8п",
            length_m=7.2,
            length_dm_raw="72",
            qty=3,
            weight=1500.0,
        ),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Chrono Client",
        db_path=db_path,
    )

    plates = _rows(db_path, "kp_plates", kp_id)
    piles = _rows(db_path, "kp_piles", kp_id)
    assert len(plates) == 2
    assert len(piles) == 1

    by_line = {
        plates[0]["line_id"]: plates[0]["position_number"],
        piles[0]["line_id"]: piles[0]["position_number"],
        plates[1]["line_id"]: plates[1]["position_number"],
    }
    assert by_line == {"ln_1": 1, "ln_2": 2, "ln_3": 3}


def test_save_mixed_piles_then_plates_keeps_order_positions(db_path: str) -> None:
    """Reverse type order still gets chronological position_number."""
    order = [
        _pile_line(line_id="ln_pile_first", mark="С120.35-12"),
        _plate_line(line_id="ln_plate_second"),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Reverse Mixed",
        db_path=db_path,
    )

    piles = _rows(db_path, "kp_piles", kp_id)
    plates = _rows(db_path, "kp_plates", kp_id)
    assert len(piles) == 1 and len(plates) == 1
    assert piles[0]["position_number"] == 1
    assert plates[0]["position_number"] == 2
    assert piles[0]["line_id"] == "ln_pile_first"
    assert plates[0]["line_id"] == "ln_plate_second"


# --- kp_meta.product_type -----------------------------------------------------


def test_save_mixed_sets_kp_meta_product_type_mixed(db_path: str) -> None:
    """≥2 distinct line product_type values → kp_meta.product_type == 'mixed'."""
    order = [_plate_line(line_id="ln_p"), _pile_line(line_id="ln_s")]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Meta Mixed",
        # Wizard may still pass last-cycle type; meta must reflect all line types.
        product_type="plates",
        db_path=db_path,
    )
    assert _meta_product_type(db_path, kp_id) == "mixed"


def test_save_single_type_plates_meta_unchanged(db_path: str) -> None:
    """Mono plates KP keeps product_type='plates' (no mixed)."""
    order = [
        _plate_line(line_id="ln_only_plate"),
        _plate_line(
            line_id="ln_only_plate_2",
            name="ПБ 72-12-8п",
            length_m=7.2,
            length_dm_raw="72",
            qty=2,
            weight=1000.0,
        ),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Mono Plates",
        product_type="plates",
        db_path=db_path,
    )
    assert _meta_product_type(db_path, kp_id) == "plates"
    assert len(_rows(db_path, "kp_plates", kp_id)) == 2
    assert len(_rows(db_path, "kp_piles", kp_id)) == 0


def test_save_single_type_piles_meta_unchanged(db_path: str) -> None:
    """Mono piles KP keeps product_type='piles' (no mixed)."""
    order = [
        _pile_line(line_id="ln_only_pile"),
        _pile_line(line_id="ln_only_pile_2", mark="С90.30-8", qty=4),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Mono Piles",
        product_type="piles",
        db_path=db_path,
    )
    assert _meta_product_type(db_path, kp_id) == "piles"
    assert len(_rows(db_path, "kp_piles", kp_id)) == 2
    assert len(_rows(db_path, "kp_plates", kp_id)) == 0


# --- line_id persistence ------------------------------------------------------


def test_save_mixed_persists_line_id_on_rows(db_path: str) -> None:
    """Create path writes order_data line_id into both kp_plates and kp_piles."""
    order = [
        _plate_line(line_id="ln_plate_abc"),
        _pile_line(line_id="ln_pile_xyz"),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Line Id Mixed",
        db_path=db_path,
    )

    plates = _rows(db_path, "kp_plates", kp_id)
    piles = _rows(db_path, "kp_piles", kp_id)
    assert plates[0]["line_id"] == "ln_plate_abc"
    assert piles[0]["line_id"] == "ln_pile_xyz"


def test_save_mono_plates_persists_line_id(db_path: str) -> None:
    """Single-type plates create also stores line_id (MNA-301 column in use)."""
    order = [_plate_line(line_id="ln_mono_plate")]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Line Id Plate",
        product_type="plates",
        db_path=db_path,
    )
    plates = _rows(db_path, "kp_plates", kp_id)
    assert len(plates) == 1
    assert plates[0]["line_id"] == "ln_mono_plate"


def test_save_mono_piles_persists_line_id(db_path: str) -> None:
    """Single-type piles create also stores line_id."""
    order = [_pile_line(line_id="ln_mono_pile")]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Line Id Pile",
        product_type="piles",
        db_path=db_path,
    )
    piles = _rows(db_path, "kp_piles", kp_id)
    assert len(piles) == 1
    assert piles[0]["line_id"] == "ln_mono_pile"


# --- MNA-303: read merge by position_number -----------------------------------


def _chrono_lines_from_kp(kp: dict) -> list[dict]:
    """Merge typed arrays by position_number (read-path contract for mixed)."""
    rows: list[dict] = []
    for key in ("plates", "piles", "steps", "marches", "bridge_piles", "fbs"):
        for row in kp.get(key) or []:
            rows.append(dict(row))
    return sorted(rows, key=lambda r: int(r.get("position_number") or 0))


def test_get_kp_by_id_mixed_loads_plates_and_piles(db_path: str) -> None:
    """Mixed meta must not fall through to plates-only: load every type table present."""
    order = [
        _plate_line(line_id="ln_a"),
        _pile_line(line_id="ln_b"),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Read Mixed Tables",
        db_path=db_path,
    )

    kp = get_kp_by_id(kp_id, db_path)
    assert kp is not None
    assert kp.get("product_type") == "mixed"
    assert len(kp.get("plates") or []) == 1
    assert len(kp.get("piles") or []) == 1
    assert (kp.get("plates") or [])[0].get("line_id") == "ln_a"
    assert (kp.get("piles") or [])[0].get("line_id") == "ln_b"


def test_get_kp_by_id_mixed_stamps_product_type_on_each_line(db_path: str) -> None:
    """Every returned line row carries product_type (plates / piles / …)."""
    order = [
        _plate_line(line_id="ln_plate"),
        _pile_line(line_id="ln_pile"),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Read Mixed Types",
        db_path=db_path,
    )

    kp = get_kp_by_id(kp_id, db_path)
    assert kp is not None
    plates = kp.get("plates") or []
    piles = kp.get("piles") or []
    assert len(plates) == 1 and len(piles) == 1
    assert plates[0].get("product_type") == "plates"
    assert piles[0].get("product_type") == "piles"


def test_get_kp_by_id_mixed_lines_chronological_by_position_number(
    db_path: str,
) -> None:
    """Merge across tables sorted by position_number matches save order."""
    order = [
        _plate_line(line_id="ln_1", name="ПБ 60-12-8п"),
        _pile_line(line_id="ln_2", mark="С120.35-12"),
        _plate_line(
            line_id="ln_3",
            name="ПБ 72-12-8п",
            length_m=7.2,
            length_dm_raw="72",
            qty=3,
            weight=1500.0,
        ),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Read Mixed Chrono",
        db_path=db_path,
    )

    kp = get_kp_by_id(kp_id, db_path)
    assert kp is not None
    chrono = _chrono_lines_from_kp(kp)
    assert [r.get("line_id") for r in chrono] == ["ln_1", "ln_2", "ln_3"]
    assert [int(r.get("position_number") or 0) for r in chrono] == [1, 2, 3]
    assert [r.get("product_type") for r in chrono] == ["plates", "piles", "plates"]


def test_get_kp_by_id_mixed_piles_then_plates_keeps_chrono_order(
    db_path: str,
) -> None:
    """Reverse type order still reads back chronologically by position_number."""
    order = [
        _pile_line(line_id="ln_pile_first"),
        _plate_line(line_id="ln_plate_second"),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Read Reverse Chrono",
        db_path=db_path,
    )

    kp = get_kp_by_id(kp_id, db_path)
    assert kp is not None
    chrono = _chrono_lines_from_kp(kp)
    assert [r.get("line_id") for r in chrono] == ["ln_pile_first", "ln_plate_second"]
    assert [r.get("product_type") for r in chrono] == ["piles", "plates"]


def test_order_data_from_kp_info_mixed_chronological_with_product_type(
    db_path: str,
) -> None:
    """Canonical KP→order_data mapper merges mixed by position_number + product_type."""
    order = [
        _plate_line(line_id="ln_1"),
        _pile_line(line_id="ln_2"),
        _plate_line(
            line_id="ln_3",
            name="ПБ 72-12-8п",
            length_m=7.2,
            length_dm_raw="72",
            qty=2,
            weight=1000.0,
        ),
    ]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        order,
        customer_name="Order Data Mixed",
        db_path=db_path,
    )

    kp = get_kp_by_id(kp_id, db_path)
    assert kp is not None
    order_data = order_data_from_kp_info(kp)

    assert len(order_data) == 3
    assert [ln.get("product_type") for ln in order_data] == [
        "plates",
        "piles",
        "plates",
    ]
    # Identity preserved when present on KP rows
    line_ids = [ln.get("line_id") for ln in order_data if ln.get("line_id")]
    if line_ids:
        assert line_ids == ["ln_1", "ln_2", "ln_3"]
    # First / third are plates (name), middle is pile (mark)
    assert "ПБ" in str(order_data[0].get("name") or "")
    assert str(order_data[1].get("mark") or order_data[1].get("name") or "").startswith(
        "С"
    )
    assert "ПБ" in str(order_data[2].get("name") or "")


# --- MNA-304: update existing kp_id (sync by line_id) + status gate ---------------


def _set_kp_status(db_path: str, kp_id: int, status: str) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET status = ? WHERE kp_id = ?",
            (status, kp_id),
        )
        conn.commit()


def _kp_status(db_path: str, kp_id: int) -> str:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT status FROM kp_meta WHERE kp_id = ?", (kp_id,))
        row = cur.fetchone()
        assert row is not None
        return str(row[0])


def test_update_kp_from_order_data_keeps_same_kp_id_on_append(db_path: str) -> None:
    """Append save must reuse kp_id (Q1=C) — no second KP_offers row."""
    initial = [_plate_line(line_id="ln_plate")]
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        initial,
        customer_name="Update Same Id",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )

    appended = [
        _plate_line(line_id="ln_plate"),
        _pile_line(line_id="ln_pile_new"),
    ]
    returned = KpPersistenceService.update_kp_from_order_data(
        kp_id,
        appended,
        db_path=db_path,
    )

    assert returned == kp_id
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM KP_offers")
        assert int(cur.fetchone()[0]) == 1
    assert len(_rows(db_path, "kp_plates", kp_id)) == 1
    assert len(_rows(db_path, "kp_piles", kp_id)) == 1
    assert _meta_product_type(db_path, kp_id) == "mixed"


def test_update_kp_from_order_data_inserts_new_line_id_row(db_path: str) -> None:
    """Unknown line_id on update → INSERT into the matching kp_* table."""
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [_plate_line(line_id="ln_a")],
        customer_name="Insert New Line",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )

    KpPersistenceService.update_kp_from_order_data(
        kp_id,
        [
            _plate_line(line_id="ln_a"),
            _pile_line(line_id="ln_b", mark="С90.30-8", qty=4),
        ],
        db_path=db_path,
    )

    piles = _rows(db_path, "kp_piles", kp_id)
    assert len(piles) == 1
    assert piles[0]["line_id"] == "ln_b"
    assert piles[0]["mark"] == "С90.30-8"
    assert piles[0]["qty"] == 4
    assert int(piles[0]["position_number"]) == 2


def test_update_kp_from_order_data_preserves_kp_plates_id_when_line_id_matches(
    db_path: str,
) -> None:
    """Matching line_id must UPDATE in place — keep kp_plates.id (production-safe)."""
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [_plate_line(line_id="ln_keep", qty=1, unit_price=1000.0)],
        customer_name="Preserve Plate Id",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )
    before = _rows(db_path, "kp_plates", kp_id)
    assert len(before) == 1
    plate_db_id = int(before[0]["id"])

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_plates SET status = ? WHERE id = ?",
            ("в производстве", plate_db_id),
        )
        conn.commit()

    KpPersistenceService.update_kp_from_order_data(
        kp_id,
        [_plate_line(line_id="ln_keep", qty=7, unit_price=2500.0)],
        db_path=db_path,
    )

    after = _rows(db_path, "kp_plates", kp_id)
    assert len(after) == 1
    assert int(after[0]["id"]) == plate_db_id
    assert after[0]["status"] == "в производстве"
    assert int(after[0]["qty"]) == 7
    assert float(after[0]["unit_price"]) == pytest.approx(2500.0)
    assert after[0]["line_id"] == "ln_keep"


def test_update_kp_from_order_data_reassigns_position_numbers_chronologically(
    db_path: str,
) -> None:
    """After append, position_number stays 1..N across tables in order_data order."""
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [_plate_line(line_id="ln_1")],
        customer_name="Update Chrono",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )

    KpPersistenceService.update_kp_from_order_data(
        kp_id,
        [
            _plate_line(line_id="ln_1"),
            _pile_line(line_id="ln_2"),
            _plate_line(
                line_id="ln_3",
                name="ПБ 72-12-8п",
                length_m=7.2,
                length_dm_raw="72",
                qty=2,
                weight=1000.0,
            ),
        ],
        db_path=db_path,
    )

    plates = _rows(db_path, "kp_plates", kp_id)
    piles = _rows(db_path, "kp_piles", kp_id)
    by_line = {
        plates[0]["line_id"]: int(plates[0]["position_number"]),
        piles[0]["line_id"]: int(piles[0]["position_number"]),
        plates[1]["line_id"]: int(plates[1]["position_number"]),
    }
    assert by_line == {"ln_1": 1, "ln_2": 2, "ln_3": 3}


@pytest.mark.parametrize(
    "blocked_status",
    ["выполнено", "отклонено", "в ожидании", "На СГП", "в работе"],
)
def test_update_kp_from_order_data_rejects_when_status_not_archived(
    db_path: str,
    blocked_status: str,
) -> None:
    """Archive-edit: update/append allowed only when kp_meta.status == «в архиве»."""
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [_plate_line(line_id="ln_gate")],
        customer_name="Status Gate",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )
    assert _kp_status(db_path, kp_id) == "в архиве"
    _set_kp_status(db_path, kp_id, blocked_status)

    with pytest.raises(ValueError, match="в архиве"):
        KpPersistenceService.update_kp_from_order_data(
            kp_id,
            [
                _plate_line(line_id="ln_gate"),
                _pile_line(line_id="ln_blocked"),
            ],
            db_path=db_path,
        )

    # No partial write: still mono plates, no pile row, status unchanged.
    assert len(_rows(db_path, "kp_plates", kp_id)) == 1
    assert len(_rows(db_path, "kp_piles", kp_id)) == 0
    assert _kp_status(db_path, kp_id) == blocked_status
    assert _meta_product_type(db_path, kp_id) == "plates"


def test_update_kp_from_order_data_allows_when_status_archived(db_path: str) -> None:
    """Happy path gate: «в архиве» permits sync; status stays «в архиве»."""
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [_plate_line(line_id="ln_ok")],
        customer_name="Status Ok",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )

    returned = KpPersistenceService.update_kp_from_order_data(
        kp_id,
        [_plate_line(line_id="ln_ok"), _pile_line(line_id="ln_ok_pile")],
        db_path=db_path,
    )
    assert returned == kp_id
    assert _kp_status(db_path, kp_id) == "в архиве"
    assert _meta_product_type(db_path, kp_id) == "mixed"


def test_offers_write_update_kp_from_order_data_delegates_same_kp_id(
    db_path: str,
) -> None:
    """Public offers_write.update_kp_from_order_data must exist and keep kp_id."""
    from core.kp import offers_write

    assert hasattr(offers_write, "update_kp_from_order_data"), (
        "offers_write.update_kp_from_order_data missing (MNA-304)"
    )

    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [_plate_line(line_id="ln_ow")],
        customer_name="Offers Write Update",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )
    returned = offers_write.update_kp_from_order_data(
        kp_id,
        [_plate_line(line_id="ln_ow"), _pile_line(line_id="ln_ow_pile")],
        db_path=db_path,
    )
    assert returned == kp_id
    assert len(_rows(db_path, "kp_piles", kp_id)) == 1


def test_update_rejects_delete_of_plate_in_production(db_path: str) -> None:
    """Omitting a plate line_id with status «в производстве» → ValueError, row kept."""
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [
            _plate_line(line_id="ln_prod"),
            _pile_line(line_id="ln_keep_pile"),
        ],
        customer_name="Protect In Production",
        product_type="mixed",
        status="в архиве",
        db_path=db_path,
    )
    plates_before = _rows(db_path, "kp_plates", kp_id)
    assert len(plates_before) == 1
    plate_db_id = int(plates_before[0]["id"])
    assert plates_before[0]["status"] == "в производстве"

    with pytest.raises(ValueError, match="в производстве"):
        KpPersistenceService.update_kp_from_order_data(
            kp_id,
            [_pile_line(line_id="ln_keep_pile")],
            db_path=db_path,
        )

    plates_after = _rows(db_path, "kp_plates", kp_id)
    assert len(plates_after) == 1
    assert int(plates_after[0]["id"]) == plate_db_id
    assert plates_after[0]["line_id"] == "ln_prod"
    assert plates_after[0]["status"] == "в производстве"
    assert len(_rows(db_path, "kp_piles", kp_id)) == 1


def test_update_rejects_when_incoming_line_missing_line_id(db_path: str) -> None:
    """Update must reject if any incoming order_data line lacks line_id."""
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.08.2026",
        [_plate_line(line_id="ln_has_id")],
        customer_name="Missing Line Id",
        product_type="plates",
        status="в архиве",
        db_path=db_path,
    )
    plates_before = _rows(db_path, "kp_plates", kp_id)
    assert len(plates_before) == 1
    plate_db_id = int(plates_before[0]["id"])

    with pytest.raises(ValueError, match="line_id"):
        KpPersistenceService.update_kp_from_order_data(
            kp_id,
            [_plate_line(line_id="ln_has_id"), _pile_line(line_id="")],
            db_path=db_path,
        )

    # No partial write: plate unchanged, no pile inserted.
    plates_after = _rows(db_path, "kp_plates", kp_id)
    assert len(plates_after) == 1
    assert int(plates_after[0]["id"]) == plate_db_id
    assert plates_after[0]["line_id"] == "ln_has_id"
    assert len(_rows(db_path, "kp_piles", kp_id)) == 0
