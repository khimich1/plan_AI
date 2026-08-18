"""ProductionSubstrateService (orch-2026-08-12-podlozhki Task 7)."""

from __future__ import annotations

from datetime import date, datetime
from unittest.mock import MagicMock, patch

import pytest

from app.services.production_substrate_service import (
    ProductionSubstrateError,
    ProductionSubstrateService,
    SubstrateRecommendation,
)
from core.production.substrate import (
    SubstrateRecommendation as CoreSubstrateRecommendation,
    extract_substrate_recommendations,
)

FIXED_NOW = datetime(2026, 8, 1, 12, 0, 0)
FIRST_FILL = date(2026, 8, 12)
DEADLINE_UNTIL = date(2026, 8, 20)


def _plate(
    *,
    plate_id: int,
    kp_id: int,
    plate_name: str,
    length_m: float = 5.7,
    width_m: float = 0.72,
    load_class: int = 800,
    qty_remaining: int = 3,
    execution_terms: str = "05.09.2026",
) -> dict:
    return {
        "plate_id": plate_id,
        "kp_id": kp_id,
        "plate_name": plate_name,
        "length_m": length_m,
        "width_m": width_m,
        "load_class": load_class,
        "qty_remaining": qty_remaining,
        "execution_terms": execution_terms,
    }


def _repo_from_backlog(backlog: list[dict]) -> MagicMock:
    """Fake KpRepository shaped like list_kps_in_production + qty_remaining."""
    by_kp: dict[int, list[dict]] = {}
    qty_by_id: dict[int, int] = {}
    terms_by_kp: dict[int, str] = {}
    for p in backlog:
        kp_id = int(p["kp_id"])
        by_kp.setdefault(kp_id, []).append(
            {
                "id": int(p["plate_id"]),
                "plate_name": p["plate_name"],
                "length_m": p["length_m"],
                "width_m": p["width_m"],
                "load_class": p["load_class"],
                "qty": int(p["qty_remaining"]),
            }
        )
        qty_by_id[int(p["plate_id"])] = int(p["qty_remaining"])
        terms_by_kp[kp_id] = str(p.get("execution_terms") or "")

    repo = MagicMock()
    repo.db_path = ":memory:"
    repo.list_kps_in_production.return_value = [
        {
            "kp_id": kp_id,
            "execution_terms": terms_by_kp.get(kp_id, ""),
            "plates": plates,
        }
        for kp_id, plates in sorted(by_kp.items())
    ]
    repo.get_plate_qty_remaining.side_effect = lambda pid: qty_by_id.get(int(pid), 0)
    repo.list_delivery_batch_items_for_in_production_plates.return_value = []
    return repo


def _opt_ok(
    *,
    primary_cuts: list[dict],
    secondary_cuts: list[dict],
) -> dict:
    return {
        "_opt_status": "ok",
        "primary_cuts": primary_cuts,
        "secondary_cuts": secondary_cuts,
    }


def _service(backlog: list[dict]) -> ProductionSubstrateService:
    return ProductionSubstrateService(kp_repository=_repo_from_backlog(backlog))


def test_substrate_recommendation_reexported_from_core() -> None:
    assert SubstrateRecommendation is CoreSubstrateRecommendation


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_cross_kp_match_one_recommendation(mock_opt: MagicMock) -> None:
    urgent = _plate(
        plate_id=10,
        kp_id=1,
        plate_name="ПБ 57-7,2",
        width_m=0.72,
        execution_terms="15.08.2026",
    )
    late = _plate(
        plate_id=20,
        kp_id=2,
        plate_name="ПБ 57-4,8",
        width_m=0.48,
        execution_terms="05.09.2026",
    )
    mock_opt.return_value = _opt_ok(
        primary_cuts=[
            {
                "primary_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": "ПБ 57-7,2",
                "rest": 480,
                "lengths": [5.7],
                "width": 720,
            }
        ],
        secondary_cuts=[
            {
                "parent_instance_id": "prim-1",
                "kp_id": 2,
                "plate_name": "ПБ 57-4,8",
                "cuts": [480],
                "qty": 1,
                "lengths": [5.7],
            }
        ],
    )

    result = _service([urgent, late]).find_substrate_recommendations(
        urgent_plate_ids=[10],
        deadline_until=DEADLINE_UNTIL,
        first_fill_target_date=FIRST_FILL,
        now=FIXED_NOW,
    )

    assert len(result) == 1
    rec = result[0]
    assert rec.plate_id == 20
    assert rec.kp_id == 2
    assert rec.plate_name == "ПБ 57-4,8"
    assert rec.qty_recommended == 1
    assert rec.under_plate_id == 10
    assert rec.under_kp_id == 1
    assert rec.under_plate_name == "ПБ 57-7,2"
    assert rec.needed_by == date(2026, 9, 5)
    assert rec.storage_days == (date(2026, 9, 5) - FIRST_FILL).days
    assert rec.saving_mm == 480
    assert rec.saving_m == pytest.approx(480 * 5.7 / 1000)
    mock_opt.assert_called_once()
    kwargs = mock_opt.call_args.kwargs
    assert "orders_2d" in kwargs
    assert len(kwargs["orders_2d"]) == 2


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_same_kp_secondary_excluded(mock_opt: MagicMock) -> None:
    a = _plate(plate_id=10, kp_id=1, plate_name="A", execution_terms="15.08.2026")
    b = _plate(plate_id=11, kp_id=1, plate_name="B", execution_terms="05.09.2026")
    mock_opt.return_value = _opt_ok(
        primary_cuts=[
            {
                "primary_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": "A",
                "rest": 400,
                "lengths": [5.0],
            }
        ],
        secondary_cuts=[
            {
                "parent_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": "B",
                "qty": 1,
            }
        ],
    )

    result = _service([a, b]).find_substrate_recommendations(
        urgent_plate_ids=[10],
        deadline_until=DEADLINE_UNTIL,
        first_fill_target_date=FIRST_FILL,
        now=FIXED_NOW,
    )
    assert result == []


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_primary_not_in_urgent_excluded(mock_opt: MagicMock) -> None:
    urgent_other = _plate(
        plate_id=99, kp_id=9, plate_name="Other", execution_terms="15.08.2026"
    )
    primary = _plate(
        plate_id=10, kp_id=1, plate_name="A", execution_terms="15.08.2026"
    )
    late = _plate(
        plate_id=20, kp_id=2, plate_name="B", execution_terms="05.09.2026"
    )
    mock_opt.return_value = _opt_ok(
        primary_cuts=[
            {
                "primary_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": "A",
                "rest": 400,
                "lengths": [5.0],
            }
        ],
        secondary_cuts=[
            {
                "parent_instance_id": "prim-1",
                "kp_id": 2,
                "plate_name": "B",
                "qty": 1,
            }
        ],
    )

    result = _service([urgent_other, primary, late]).find_substrate_recommendations(
        urgent_plate_ids=[99],  # primary 10 not urgent
        deadline_until=DEADLINE_UNTIL,
        first_fill_target_date=FIRST_FILL,
        now=FIXED_NOW,
    )
    assert result == []


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_qty_recommended_counts_multiple_secondaries(mock_opt: MagicMock) -> None:
    urgent = _plate(
        plate_id=10, kp_id=1, plate_name="A", execution_terms="15.08.2026"
    )
    late = _plate(
        plate_id=20, kp_id=2, plate_name="B", execution_terms="05.09.2026"
    )
    mock_opt.return_value = _opt_ok(
        primary_cuts=[
            {
                "primary_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": "A",
                "rest": 300,
                "lengths": [6.0],
            },
            {
                "primary_instance_id": "prim-2",
                "kp_id": 1,
                "plate_name": "A",
                "rest": 500,
                "lengths": [6.0],
            },
        ],
        secondary_cuts=[
            {
                "parent_instance_id": "prim-1",
                "kp_id": 2,
                "plate_name": "B",
                "qty": 1,
            },
            {
                "parent_instance_id": "prim-2",
                "kp_id": 2,
                "plate_name": "B",
                "qty": 1,
            },
            {
                "parent_instance_id": "prim-2",
                "kp_id": 2,
                "plate_name": "B",
                "qty": 1,
            },
        ],
    )

    result = _service([urgent, late]).find_substrate_recommendations(
        urgent_plate_ids=[10],
        deadline_until=DEADLINE_UNTIL,
        first_fill_target_date=FIRST_FILL,
        now=FIXED_NOW,
    )

    assert len(result) == 1
    assert result[0].qty_recommended == 3
    # Aggregation keeps max saving_mm among cuts (500 > 300)
    assert result[0].saving_mm == 500
    assert result[0].saving_m == pytest.approx(500 * 6.0 / 1000)


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_storage_days_and_saving_m_math(mock_opt: MagicMock) -> None:
    urgent = _plate(
        plate_id=10,
        kp_id=1,
        plate_name="Urgent",
        length_m=5.2,
        execution_terms="15.08.2026",
    )
    late = _plate(
        plate_id=20,
        kp_id=2,
        plate_name="Late",
        length_m=5.2,
        execution_terms="01.09.2026",
    )
    rest_mm = 480
    length_m = 5.2
    mock_opt.return_value = _opt_ok(
        primary_cuts=[
            {
                "primary_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": "Urgent",
                "rest": rest_mm,
                "lengths": [length_m],
            }
        ],
        secondary_cuts=[
            {
                "parent_instance_id": "prim-1",
                "kp_id": 2,
                "plate_name": "Late",
                "qty": 1,
            }
        ],
    )

    first_fill = date(2026, 8, 10)
    result = _service([urgent, late]).find_substrate_recommendations(
        urgent_plate_ids=[10],
        deadline_until=DEADLINE_UNTIL,
        first_fill_target_date=first_fill,
        now=FIXED_NOW,
    )

    assert len(result) == 1
    rec = result[0]
    assert rec.needed_by == date(2026, 9, 1)
    assert rec.storage_days == (date(2026, 9, 1) - first_fill).days
    assert rec.storage_days == 22
    assert rec.saving_mm == rest_mm
    assert rec.saving_m == pytest.approx(rest_mm * length_m / 1000)


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_empty_backlog_returns_empty(mock_opt: MagicMock) -> None:
    result = _service([]).find_substrate_recommendations(
        urgent_plate_ids=[1],
        deadline_until=DEADLINE_UNTIL,
        first_fill_target_date=FIRST_FILL,
        now=FIXED_NOW,
    )
    assert result == []
    mock_opt.assert_not_called()


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_optimizer_error_raises_domain_error(mock_opt: MagicMock) -> None:
    urgent = _plate(plate_id=10, kp_id=1, plate_name="A")
    late = _plate(plate_id=20, kp_id=2, plate_name="B")
    mock_opt.return_value = {
        "_opt_status": "error",
        "_opt_error_code": "solver_infeasible",
        "_opt_error_message": "infeasible",
    }

    with pytest.raises(ProductionSubstrateError, match="infeasible"):
        _service([urgent, late]).find_substrate_recommendations(
            urgent_plate_ids=[10],
            deadline_until=DEADLINE_UNTIL,
            first_fill_target_date=FIRST_FILL,
            now=FIXED_NOW,
        )


@patch(
    "app.services.production_substrate_service.optimize_with_cascading_longitudinal_cuts"
)
def test_needed_by_uses_produce_by_over_execution_terms(mock_opt: MagicMock) -> None:
    """AC: needed_by = deadline of late plate (produce_by or execution_terms)."""
    urgent = _plate(
        plate_id=10,
        kp_id=1,
        plate_name="Urgent",
        execution_terms="15.08.2026",
    )
    late = _plate(
        plate_id=20,
        kp_id=2,
        plate_name="Late",
        execution_terms="05.09.2026",  # later than produce_by
    )
    mock_opt.return_value = _opt_ok(
        primary_cuts=[
            {
                "primary_instance_id": "prim-1",
                "kp_id": 1,
                "plate_name": "Urgent",
                "rest": 400,
                "lengths": [5.7],
            }
        ],
        secondary_cuts=[
            {
                "parent_instance_id": "prim-1",
                "kp_id": 2,
                "plate_name": "Late",
                "qty": 1,
            }
        ],
    )
    svc = _service([urgent, late])
    svc.kp_repository.list_delivery_batch_items_for_in_production_plates.return_value = [
        {
            "plate_id": 20,
            "produce_by": "2026-08-25",
            "qty": 3,
            "batch_name": "Этаж 2",
        }
    ]

    result = svc.find_substrate_recommendations(
        urgent_plate_ids=[10],
        deadline_until=DEADLINE_UNTIL,
        first_fill_target_date=FIRST_FILL,
        now=FIXED_NOW,
    )

    assert len(result) == 1
    assert result[0].needed_by == date(2026, 8, 25)
    assert result[0].storage_days == (date(2026, 8, 25) - FIRST_FILL).days


def test_extract_skips_orphan_secondary_and_missing_deadline() -> None:
    plate_map = {(1, "A"): 10, (2, "B"): 20, (3, "C"): 30}
    result = extract_substrate_recommendations(
        {
            "primary_cuts": [
                {
                    "primary_instance_id": "prim-1",
                    "kp_id": 1,
                    "plate_name": "A",
                    "rest": 400,
                    "lengths": [5.0],
                }
            ],
            "secondary_cuts": [
                {
                    "parent_instance_id": "missing",
                    "kp_id": 2,
                    "plate_name": "B",
                    "qty": 1,
                },
                {
                    "parent_instance_id": "prim-1",
                    "kp_id": 3,
                    "plate_name": "C",
                    "qty": 1,
                },
            ],
        },
        plate_id_by_kp_name=plate_map,
        urgent_plate_ids=[10],
        deadline_by_plate_id={20: date(2026, 9, 1)},  # C has no deadline
        first_fill_target_date=FIRST_FILL,
    )
    assert result == []


def test_extract_sorts_by_saving_m_descending() -> None:
    plate_map = {(1, "A"): 10, (2, "B"): 20, (3, "C"): 30}
    result = extract_substrate_recommendations(
        {
            "primary_cuts": [
                {
                    "primary_instance_id": "prim-small",
                    "kp_id": 1,
                    "plate_name": "A",
                    "rest": 200,
                    "lengths": [5.0],
                },
                {
                    "primary_instance_id": "prim-large",
                    "kp_id": 1,
                    "plate_name": "A",
                    "rest": 500,
                    "lengths": [5.0],
                },
            ],
            "secondary_cuts": [
                {
                    "parent_instance_id": "prim-small",
                    "kp_id": 2,
                    "plate_name": "B",
                    "qty": 1,
                },
                {
                    "parent_instance_id": "prim-large",
                    "kp_id": 3,
                    "plate_name": "C",
                    "qty": 1,
                },
            ],
        },
        plate_id_by_kp_name=plate_map,
        urgent_plate_ids=[10],
        deadline_by_plate_id={
            20: date(2026, 9, 1),
            30: date(2026, 9, 5),
        },
        first_fill_target_date=FIRST_FILL,
    )

    assert [r.plate_id for r in result] == [30, 20]
    assert result[0].saving_mm == 500
    assert result[1].saving_mm == 200
