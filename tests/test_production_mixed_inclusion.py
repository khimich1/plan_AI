"""MNA-701 (TDD RED): mixed-with-plates KPs appear in production candidates.

Acceptance:
- ``mixed`` KP that has ``kp_plates`` rows appears in production candidates
- Non-plate-only KP (e.g. piles-only) remains excluded
- Candidate ``plates`` come only from ``kp_plates`` (not piles / other tables)
"""

from __future__ import annotations

import sqlite3

import pytest

from app.repositories.kp_repository import KpRepository
from app.services.production_service import ProductionService
from core.kp_db_schema import init_schema
from core.kp_persistence_service import KpPersistenceService


@pytest.fixture()
def db_path(tmp_path) -> str:
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
        "unit_price": 100.0,
    }
    base.update(overrides)
    return base


def _save_mixed_plates_piles(db_path: str, *, customer_name: str = "Mixed KP") -> int:
    return KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        [
            _plate_line(line_id="ln_a", name="ПБ 60-12-8п"),
            _pile_line(line_id="ln_b", mark="С120.35-12"),
        ],
        customer_name=customer_name,
        status="в работе",
        db_path=db_path,
    )


def _save_piles_only(db_path: str, *, customer_name: str = "Pile-only KP") -> int:
    return KpPersistenceService.save_kp_to_db(
        "02.01.2026",
        [_pile_line(line_id="ln_pile_only", mark="С120.35-12")],
        customer_name=customer_name,
        status="в работе",
        product_type="piles",
        db_path=db_path,
    )


def _save_plates_only(db_path: str, *, customer_name: str = "Plate KP") -> int:
    return KpPersistenceService.save_kp_to_db(
        "03.01.2026",
        [_plate_line(line_id="ln_plate_only", name="ПБ 60-12-8п")],
        customer_name=customer_name,
        status="в работе",
        db_path=db_path,
    )


def _set_kp_plates_in_production(db_path: str, kp_id: int) -> list[int]:
    """Mark all kp_plates for kp_id as in production; return their ids."""
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE kp_plates SET status = 'в производстве' WHERE kp_id = ?",
            (kp_id,),
        )
        conn.commit()
        cur.execute(
            "SELECT id FROM kp_plates WHERE kp_id = ? ORDER BY id",
            (kp_id,),
        )
        return [int(r[0]) for r in cur.fetchall()]


def test_list_kps_in_production_includes_mixed_with_plates(db_path: str) -> None:
    """mixed KP with kp_plates rows must appear alongside mono plates KP."""
    mixed_id = _save_mixed_plates_piles(db_path, customer_name="Mixed KP")
    plates_id = _save_plates_only(db_path, customer_name="Plate KP")

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = ?", (mixed_id,))
        assert cur.fetchone()[0] == "mixed"

    repo = KpRepository(db_path=db_path)
    items = repo.list_kps_in_production()
    by_id = {int(item["kp_id"]): item for item in items}

    assert mixed_id in by_id, "mixed-with-plates must be a production candidate KP"
    assert plates_id in by_id
    assert by_id[mixed_id]["customer_name"] == "Mixed KP"


def test_list_kps_in_production_excludes_non_plate_only_kp(db_path: str) -> None:
    """Piles-only (no kp_plates) stays excluded even when mixed-with-plates is included."""
    mixed_id = _save_mixed_plates_piles(db_path, customer_name="Mixed KP")
    piles_id = _save_piles_only(db_path, customer_name="Pile-only KP")

    repo = KpRepository(db_path=db_path)
    items = repo.list_kps_in_production()
    kp_ids = {int(item["kp_id"]) for item in items}

    assert mixed_id in kp_ids
    assert piles_id not in kp_ids
    assert all(item["customer_name"] != "Pile-only KP" for item in items)


def test_list_kp_candidates_mixed_uses_only_kp_plates_rows(db_path: str) -> None:
    """Production wizard candidates for mixed expose only kp_plates rows, not piles."""
    mixed_id = _save_mixed_plates_piles(db_path, customer_name="Mixed KP")
    plate_ids = _set_kp_plates_in_production(db_path, mixed_id)
    assert len(plate_ids) == 1

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = ?", (mixed_id,))
        assert cur.fetchone()[0] == 1

    repo = KpRepository(db_path=db_path)
    service = ProductionService(kp_repository=repo)
    candidates = service.list_kp_candidates()

    assert candidates["count"] == 1
    item = candidates["items"][0]
    assert item["kp_id"] == mixed_id

    plates = item["plates"]
    assert len(plates) == 1
    assert plates[0]["id"] == plate_ids[0]
    assert plates[0]["plate_name"] == "ПБ 60-12-8п"
    # Must not surface pile marks as production plate rows.
    assert all(p.get("plate_name") != "С120.35-12" for p in plates)
