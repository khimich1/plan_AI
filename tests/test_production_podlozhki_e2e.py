"""E2E: analyze-substrates → plans/build (orch-2026-08-12-podlozhki Task 13)."""

from __future__ import annotations

import sqlite3
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.domain.enums import PlateStatus
from app.services.production_planning_service import ProductionPlanningService
from tests.helpers import production_api_fixtures as paf

API_PREFIX = paf.API_PREFIX
DATE_KEY = paf.DATE_KEY
_TS = "2026-04-01T12:00:00"

FILL_TARGETS = [
    {"date": "2026-04-20", "tracks": 5},
    {"date": "2026-04-21", "tracks": 3},
]
DEADLINE_UNTIL = "2026-04-25"
URGENT_NAME = "ПБ 57-7,2"
LATE_NAME = "ПБ 57-4,8"


def _seed_delivery_batch(
    db_path: str,
    *,
    kp_id: int,
    plate_id: int,
    produce_by: str,
    qty: int,
    batch_name: str = "П1",
) -> None:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO delivery_schedule (
                kp_id, invoice_number, contract_number, status, created_at, updated_at
            ) VALUES (?, 'СЧ-E2E', 'Д-E2E', 'draft', ?, ?)
            """,
            (kp_id, _TS, _TS),
        )
        schedule_id = int(cur.lastrowid)
        cur.execute(
            """
            INSERT INTO delivery_batch (
                schedule_id, name, deliver_from, deliver_to, produce_by, sort_order
            ) VALUES (?, ?, '2026-04-25', '2026-04-30', ?, 1)
            """,
            (schedule_id, batch_name, produce_by),
        )
        batch_id = int(cur.lastrowid)
        cur.execute(
            "INSERT INTO delivery_batch_item (batch_id, plate_id, qty) VALUES (?, ?, ?)",
            (batch_id, plate_id, qty),
        )
        conn.commit()


def _seed_podlozhki_scenario(db_path: str) -> tuple[int, int]:
    """Urgent KP1 + late KP2 geometry (same shape as analyze happy-path)."""
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_plates SET plate_name = ?, length_m = ?, width_m = ? WHERE kp_id = 1",
            (URGENT_NAME, 5.7, 0.72),
        )
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (2, '2026-01-02', '05.09.2026', 'ПозднийКлиент')"
        )
        conn.execute("INSERT INTO kp_meta (kp_id, status) VALUES (2, 'в работе')")
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (2, 1, ?, 5.7, 0.48, 800, 2, ?)
            """,
            (LATE_NAME, PlateStatus.IN_PRODUCTION.value),
        )
        conn.commit()

        urgent_plate_id = int(
            conn.execute("SELECT id FROM kp_plates WHERE kp_id = 1").fetchone()[0]
        )
        late_plate_id = int(
            conn.execute("SELECT id FROM kp_plates WHERE kp_id = 2").fetchone()[0]
        )
        urgent_qty = int(
            conn.execute("SELECT qty FROM kp_plates WHERE id = ?", (urgent_plate_id,)).fetchone()[0]
        )

    _seed_delivery_batch(
        db_path,
        kp_id=1,
        plate_id=urgent_plate_id,
        produce_by="2026-04-22",
        qty=urgent_qty,
        batch_name="1 этаж",
    )
    return urgent_plate_id, late_plate_id


def _plate_rows(db_path: str, plate_ids: list[int]) -> list[tuple[int, str | None, str]]:
    placeholders = ",".join("?" * len(plate_ids))
    with sqlite3.connect(db_path) as conn:
        rows = conn.execute(
            f"SELECT id, plan_id, status FROM kp_plates WHERE id IN ({placeholders}) ORDER BY id",
            plate_ids,
        ).fetchall()
    return [(int(r[0]), r[1], str(r[2])) for r in rows]


def _fake_analyze_optimize(*, orders_2d, **kwargs):
    return {
        "_opt_status": "ok",
        "primary_cuts": [
            {
                "primary_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": URGENT_NAME,
                "rest": 480,
                "lengths": [5.7],
                "width": 720,
            }
        ],
        "secondary_cuts": [
            {
                "parent_instance_id": "prim-1",
                "kp_id": 2,
                "plate_name": LATE_NAME,
                "cuts": [480],
                "qty": 1,
                "lengths": [5.7],
            }
        ],
    }


def _fake_build_optimize_all(self, *, orders_2d, **kwargs):
    """One track covering every selected order (unlike default first-order-only fake)."""
    if not orders_2d:
        return [], {}
    items: list[dict[str, Any]] = []
    assignments: list[dict[str, Any]] = []
    for order in orders_2d:
        for _ in range(int(order.get("qty") or 0)):
            item = {
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"],
                "load_code": order["load_code"],
            }
            items.append(item)
            assignments.append({**item, "source": "primary"})
    return (
        [{"label": "ОСНОВНАЯ", "items": items}],
        {"total_plates": len(items), "plate_assignments": assignments},
    )


def _put_day_capacity(
    client: TestClient,
    cookies: dict[str, str],
    *,
    day: str,
    max_tracks: int,
) -> None:
    response = client.put(
        f"{API_PREFIX}/day-capacity",
        json={"date": day, "max_tracks": max_tracks},
        cookies=cookies,
    )
    assert response.status_code == 200, response.text


def test_podlozhki_happy_path(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_api_db: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    urgent_plate_id, late_plate_id = _seed_podlozhki_scenario(production_api_db)

    for day, tracks in (("2026-04-20", 5), ("2026-04-21", 3)):
        _put_day_capacity(
            production_api_client,
            production_admin_cookie,
            day=day,
            max_tracks=tracks,
        )

    monkeypatch.setattr(
        "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts",
        _fake_analyze_optimize,
    )
    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        _fake_build_optimize_all,
    )

    analyze = production_api_client.post(
        f"{API_PREFIX}/analyze-substrates",
        json={"fill_targets": FILL_TARGETS, "deadline_until": DEADLINE_UNTIL},
        cookies=production_admin_cookie,
    )
    assert analyze.status_code == 200, analyze.text
    payload = analyze.json()
    assert set(payload) >= {
        "urgent_positions",
        "substrate_recommendations",
        "capacity_deficit",
        "analysis_meta",
    }
    urgent_ids = {item["plate_id"] for item in payload["urgent_positions"]}
    assert urgent_plate_id in urgent_ids
    assert payload["analysis_meta"]["optimization_status"] == "ok"
    assert len(payload["substrate_recommendations"]) >= 1
    assert payload["substrate_recommendations"][0]["plate_id"] == late_plate_id

    selected_plate_ids = {
        "1": [urgent_plate_id],
        "2": [late_plate_id],
    }
    build = production_api_client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 3,
            "filter_method": "kp",
            "selected_kp_ids": [1, 2],
            "selected_plate_ids": selected_plate_ids,
            "fill_targets": FILL_TARGETS,
        },
        cookies=production_admin_cookie,
    )
    assert build.status_code == 200, build.text
    plan_id = build.json()["plan"]["id"]
    assert plan_id

    for plate_id, row_plan_id, status in _plate_rows(
        production_api_db, [urgent_plate_id, late_plate_id]
    ):
        assert row_plan_id == plan_id, f"plate {plate_id} plan_id={row_plan_id}"
        assert status == PlateStatus.IN_PLAN.value, f"plate {plate_id} status={status}"


def test_build_failure_leaves_plates_unplanned(
    production_api_client: TestClient,
    production_admin_cookie: dict[str, str],
    production_api_db: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    urgent_plate_id, late_plate_id = _seed_podlozhki_scenario(production_api_db)
    plate_ids = [urgent_plate_id, late_plate_id]

    monkeypatch.setattr(
        ProductionPlanningService,
        "_run_optimization_and_split",
        lambda self, *, orders_2d, **kwargs: ([], {}),
    )

    response = production_api_client.post(
        f"{API_PREFIX}/plans/build",
        json={
            "start_date": DATE_KEY,
            "tracks_count": 1,
            "filter_method": "kp",
            "selected_kp_ids": [1, 2],
            "selected_plate_ids": {
                "1": [urgent_plate_id],
                "2": [late_plate_id],
            },
        },
        cookies=production_admin_cookie,
    )
    assert response.status_code == 422, response.text
    assert isinstance(response.json()["detail"], str)

    for plate_id, plan_id, status in _plate_rows(production_api_db, plate_ids):
        assert plan_id is None, f"plate {plate_id} unexpectedly planned: {plan_id}"
        assert status == PlateStatus.IN_PRODUCTION.value, (
            f"plate {plate_id} status={status}"
        )
