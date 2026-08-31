"""MARCH-603: march KPs excluded from production wizard candidates."""

from __future__ import annotations

import sqlite3

import pytest

from app.repositories.kp_repository import KpRepository
from core.kp_db_schema import init_schema
from core.kp_persistence_service import KpPersistenceService


@pytest.fixture()
def db_path(tmp_path) -> str:
    path = str(tmp_path / "plita.db")
    init_schema(path)
    return path


def test_list_kps_in_production_excludes_marches(db_path: str) -> None:
    KpPersistenceService.save_kp_to_db(
        "01.01.2026",
        [
            {
                "product_kind": "march",
                "name": "1ЛМ 27-11-14-4",
                "mark": "1ЛМ 27-11-14-4",
                "concrete_grade": "B25",
                "qty": 2,
                "unit_price": 100.0,
            }
        ],
        customer_name="March KP",
        status="в работе",
        product_type="marches",
        db_path=db_path,
    )
    KpPersistenceService.save_kp_to_db(
        "02.01.2026",
        [
            {
                "name": "ПБ 60-12-8п",
                "length_m": 6.0,
                "width_m": 1.2,
                "load_class": 800,
                "qty": 1,
                "unit_price": 1000.0,
                "weight": 500.0,
            }
        ],
        customer_name="Plate KP",
        status="в работе",
        db_path=db_path,
    )

    repo = KpRepository(db_path=db_path)
    items = repo.list_kps_in_production()
    assert len(items) == 1
    assert items[0]["customer_name"] == "Plate KP"

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE kp_plates SET status = 'в производстве' WHERE kp_id = 2"
        )
        conn.commit()

    from app.services.production_service import ProductionService

    service = ProductionService(kp_repository=repo)
    candidates = service.list_kp_candidates()
    assert candidates["count"] == 1
    assert candidates["items"][0]["kp_id"] == 2
