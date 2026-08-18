"""Unit tests for core/production/planning pipeline phases."""

from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
from core.production.capacity import validate_fill_targets
from core.production.dto import (
    LoadConfig,
    OptimizeConfig,
    PersistConfig,
    PlanBuildInput,
)
from core.production.errors import PlanBuildError
from core.production.planning import (
    build_tracks_by_day_from_targets,
    load,
    optimize,
    persist,
    trim_assignments_to_tracks,
    validate,
)

PLATE_NAME = "ПБ 60-12-8п"


def _plan_load_for_db(db_path: str):
    from app.repositories.plan_repository import PlanRepository
    from app.services.plan_distribution_service import PlanLoadAdapter

    return PlanLoadAdapter(PlanRepository(db_path=db_path))


@pytest.fixture
def tmp_plita(tmp_path) -> str:
    db_path = str(tmp_path / "plita.db")
    kp_db.init_schema(db_path)

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (1, '2026-01-01', '21.04.2026', 'ТестКлиент')"
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (1, 'в работе')"
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, length_m, width_m,
                load_class, qty, status
            ) VALUES (1, 1, ?, 6.0, 1.2, 800, 3, 'в производстве')
            """,
            (PLATE_NAME,),
        )
        conn.commit()
    return db_path


@pytest.fixture(autouse=True)
def _patch_reinforcement(monkeypatch):
    monkeypatch.setattr(
        "core.production.planning.get_reinforcement",
        lambda **kwargs: 999.0,
    )


# ----- validate -----


def test_validate_rejects_bad_start_date() -> None:
    plan_input = PlanBuildInput(
        start_date="not-a-date",
        tracks_count=3,
        filter_method="all",
    )
    with pytest.raises(PlanBuildError, match="start_date"):
        validate(plan_input)


def test_validate_rejects_tracks_out_of_range() -> None:
    with pytest.raises(PlanBuildError, match="tracks_count"):
        validate(
            PlanBuildInput(
                start_date="2026-04-21",
                tracks_count=0,
                filter_method="all",
            )
        )
    with pytest.raises(PlanBuildError, match="tracks_count"):
        validate(
            PlanBuildInput(
                start_date="2026-04-21",
                tracks_count=51,
                filter_method="all",
            )
        )


def test_validate_rejects_kp_filter_without_ids() -> None:
    with pytest.raises(PlanBuildError, match="selected_kp_ids"):
        validate(
            PlanBuildInput(
                start_date="2026-04-21",
                tracks_count=3,
                filter_method="kp",
                selected_kp_ids=(),
            )
        )


def test_validate_accepts_valid_input() -> None:
    validate(
        PlanBuildInput(
            start_date="2026-04-21",
            tracks_count=3,
            filter_method="all",
        )
    )


# ----- load -----


def test_load_raises_when_no_kps(tmp_path) -> None:
    db_path = str(tmp_path / "empty.db")
    kp_db.init_schema(db_path)

    with pytest.raises(PlanBuildError, match="Нет подходящих КП"):
        load(
            PlanBuildInput(
                start_date="2026-04-21",
                tracks_count=3,
                filter_method="all",
            ),
            config=LoadConfig(plita_db_path=db_path, pb_db_path=db_path),
            plan_load=_plan_load_for_db(db_path),
        )


def test_load_returns_plates_and_orders(tmp_plita) -> None:
    result = load(
        PlanBuildInput(
            start_date="2026-04-21",
            tracks_count=3,
            filter_method="all",
        ),
        config=LoadConfig(plita_db_path=tmp_plita, pb_db_path=tmp_plita),
        plan_load=_plan_load_for_db(tmp_plita),
    )

    assert len(result.kp_list) == 1
    assert result.kp_list[0]["kp_id"] == 1
    assert len(result.selected_plates) == 1
    assert result.selected_plates[0]["qty"] == 3
    assert len(result.orders_2d) == 1
    assert result.orders_2d[0]["qty"] == 3
    assert result.orders_2d[0]["plate_name"] == PLATE_NAME
    assert (6.0, 1200) in result.plate_lookup_exact
    assert 6.0 in result.plate_lookup_by_length


def test_load_raises_qty_above_available(tmp_plita) -> None:
    with sqlite3.connect(tmp_plita) as conn:
        plate_id = conn.execute(
            "SELECT id FROM kp_plates WHERE kp_id = 1"
        ).fetchone()[0]

    with pytest.raises(PlanBuildError, match="доступно 3"):
        load(
            PlanBuildInput(
                start_date="2026-04-21",
                tracks_count=3,
                filter_method="kp",
                selected_kp_ids=(1,),
                selected_plate_ids={1: [plate_id]},
                selected_plate_qty={1: {plate_id: 10}},
            ),
            config=LoadConfig(plita_db_path=tmp_plita, pb_db_path=tmp_plita),
            plan_load=_plan_load_for_db(tmp_plita),
        )


def test_load_partial_qty(tmp_plita) -> None:
    with sqlite3.connect(tmp_plita) as conn:
        plate_id = conn.execute(
            "SELECT id FROM kp_plates WHERE kp_id = 1"
        ).fetchone()[0]

    result = load(
        PlanBuildInput(
            start_date="2026-04-21",
            tracks_count=3,
            filter_method="kp",
            selected_kp_ids=(1,),
            selected_plate_ids={1: [plate_id]},
            selected_plate_qty={1: {plate_id: 2}},
        ),
        config=LoadConfig(plita_db_path=tmp_plita, pb_db_path=tmp_plita),
        plan_load=_plan_load_for_db(tmp_plita),
    )
    assert result.orders_2d[0]["qty"] == 2


# ----- optimize -----


def test_optimize_returns_empty_when_no_orders() -> None:
    from core.production.dto import LoadResult

    result = optimize(
        LoadResult(
            kp_list=[],
            selected_plates=[],
            orders_2d=[],
            plate_lookup_exact={},
            plate_lookup_by_length={},
        ),
        config=OptimizeConfig(pb_db_path="unused"),
    )
    assert result.all_tracks_list == []
    assert result.optimization_result == {}


def test_optimize_with_mocked_optimizer(monkeypatch, tmp_plita) -> None:
    load_result = load(
        PlanBuildInput(
            start_date="2026-04-21",
            tracks_count=3,
            filter_method="all",
        ),
        config=LoadConfig(plita_db_path=tmp_plita, pb_db_path=tmp_plita),
        plan_load=_plan_load_for_db(tmp_plita),
    )
    order = load_result.orders_2d[0]

    def fake_optimize(*, orders_2d, **kwargs):
        qty = int(orders_2d[0]["qty"])
        return {
            "total_plates": qty,
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                    "unit_id": f"u{i}",
                }
                for i in range(qty)
            ],
        }

    def fake_build_layout_sequence(*, runtime, **kwargs):
        return [{"sequence": [{"kp_id": order["kp_id"], "plate_name": order["plate_name"]}]}]

    def fake_split(*seq, **kwargs):
        items = [
            {
                "kp_id": order["kp_id"],
                "plate_name": order["plate_name"],
                "length": order["length"],
                "width": order["width"] / 1000.0,
                "load_code": order["load_code"],
                "unit_id": f"u{i}",
            }
            for i in range(int(order["qty"]))
        ]
        return [{"label": "ОСНОВНАЯ", "items": items}]

    monkeypatch.setattr(
        "core.production.planning.optimize_with_cascading_longitudinal_cuts",
        fake_optimize,
    )
    monkeypatch.setattr(
        "core.production.planning.build_layout_sequence",
        fake_build_layout_sequence,
    )
    monkeypatch.setattr(
        "core.production.planning.split_sequence_into_tracks",
        fake_split,
    )
    monkeypatch.setattr(
        "core.production.planning.build_rescue_tracks",
        lambda **kwargs: ([], {}, []),
    )

    result = optimize(
        load_result,
        config=OptimizeConfig(pb_db_path=tmp_plita, track_top_up_from_following=False),
    )

    assert len(result.all_tracks_list) == 1
    assert len(result.all_tracks_list[0]["items"]) == 3
    assert result.optimization_result["total_plates"] == 3


def test_optimize_pipeline_validate_load_optimize(monkeypatch, tmp_plita) -> None:
    """End-to-end through validate → load → optimize with mocked heavy steps."""
    plan_input = PlanBuildInput(
        start_date="2026-04-21",
        tracks_count=3,
        filter_method="all",
    )
    validate(plan_input)
    load_result = load(
        plan_input,
        config=LoadConfig(plita_db_path=tmp_plita, pb_db_path=tmp_plita),
        plan_load=_plan_load_for_db(tmp_plita),
    )
    order = load_result.orders_2d[0]

    monkeypatch.setattr(
        "core.production.planning.optimize_with_cascading_longitudinal_cuts",
        lambda *, orders_2d, **kw: {
            "total_plates": orders_2d[0]["qty"],
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
            ],
        },
    )
    monkeypatch.setattr(
        "core.production.planning.build_layout_sequence",
        lambda **kw: [],
    )
    monkeypatch.setattr(
        "core.production.planning.split_sequence_into_tracks",
        lambda seq, **kw: [{"label": "T", "items": [{"kp_id": 1, "plate_name": PLATE_NAME}]}],
    )
    monkeypatch.setattr(
        "core.production.planning.build_rescue_tracks",
        lambda **kwargs: ([], {}, []),
    )

    opt_result = optimize(
        load_result,
        config=OptimizeConfig(pb_db_path=tmp_plita),
    )
    assert opt_result.all_tracks_list
    assert opt_result.optimization_result.get("total_plates") == 3


# ----- fill_targets helpers -----


def test_validate_fill_targets_rejects_over_capacity() -> None:
    with pytest.raises(PlanBuildError, match="свободно"):
        validate_fill_targets(
            [{"date": "2026-04-27", "tracks": 3}],
            {"2026-04-27": 5},
            occupancy={"2026-04-27": 4},
        )


def test_validate_fill_targets_uses_day_capacity_override() -> None:
    """Override day_max=3 → free=3−1=2; requesting 3 fails."""
    with pytest.raises(PlanBuildError, match="свободно 2"):
        validate_fill_targets(
            [{"date": "2026-04-27", "tracks": 3}],
            {"2026-04-27": 3},
            occupancy={"2026-04-27": 1},
        )


def test_validate_fill_targets_clamps_day_capacity_above_hard_cap() -> None:
    """Stale override 9 is clamped to 5; free=5−0=5; requesting 6 fails."""
    with pytest.raises(PlanBuildError, match="свободно 5"):
        validate_fill_targets(
            [{"date": "2026-04-27", "tracks": 6}],
            {"2026-04-27": 9},
        )


def test_validate_fill_targets_occupancy_free_slots() -> None:
    """occupancy=3, max=5 → free=2; tracks=2 OK; tracks=4 → PlanBuildError."""
    day_capacity = {"2026-04-27": 5}
    occupancy = {"2026-04-27": 3}
    validate_fill_targets(
        [{"date": "2026-04-27", "tracks": 2}],
        day_capacity,
        occupancy=occupancy,
    )
    with pytest.raises(PlanBuildError, match=r"свободно 2.*запрошено 4"):
        validate_fill_targets(
            [{"date": "2026-04-27", "tracks": 4}],
            day_capacity,
            occupancy=occupancy,
        )


def test_build_tracks_by_day_from_targets_splits() -> None:
    tracks = [{"label": f"T{i}"} for i in range(5)]
    by_day = build_tracks_by_day_from_targets(
        kept_tracks=tracks,
        fill_targets=[
            {"date": "2026-04-27", "tracks": 2},
            {"date": "2026-04-28", "tracks": 3},
        ],
    )
    assert len(by_day["2026-04-27"]) == 2
    assert len(by_day["2026-04-28"]) == 3


def test_trim_assignments_to_tracks_keeps_counts() -> None:
    kept = [
        {
            "label": "ОСНОВНАЯ",
            "items": [{"kp_id": 1, "plate_name": "ПБ"}],
        }
    ]
    opt = {
        "plate_assignments": [
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ"},
            {"source": "primary", "kp_id": 1, "plate_name": "ПБ"},
        ]
    }
    trimmed = trim_assignments_to_tracks(
        optimization_result=opt,
        kept_tracks=kept,
    )
    assert len(trimmed["plate_assignments"]) == 1


# ----- persist -----


class _FakePersistRepo:
    def __init__(self) -> None:
        self.plans: dict[str, dict] = {}
        self.active_id: str | None = None
        self.occupancy: dict[str, int] = {}

    def get_global_occupancy(self, *, exclude_plan_id=None):
        return dict(self.occupancy)

    def build_plan_from_tracks(self, **kwargs):
        plan_id = kwargs.get("plan_id") or "plan_test_1"
        days: dict[str, dict] = {}
        if kwargs.get("precomputed_tracks_by_day"):
            for i, (date_key, tracks) in enumerate(
                sorted(kwargs["precomputed_tracks_by_day"].items()), start=1
            ):
                days[date_key] = {"day_number": i, "tracks": tracks}
        else:
            days[kwargs["start_date"]] = {
                "day_number": 1,
                "tracks": kwargs["new_tracks_list"],
            }
        plan = {
            "id": plan_id,
            "start_date": kwargs["start_date"],
            "tracks_count": kwargs["tracks_per_day"],
            "days": days,
        }
        return plan, {"is_new_plan": True}

    def get(self, plan_id):
        if plan_id not in self.plans:
            return None
        return {"payload": self.plans[plan_id], "version": 1}

    def create(self, payload):
        self.plans[payload["id"]] = payload
        return {"payload": payload, "version": 1}

    def save(self, payload, expected_version):
        self.plans[payload["id"]] = payload
        return {"payload": payload, "version": expected_version + 1}

    def set_active(self, plan_id):
        self.active_id = plan_id


def test_persist_saves_plan_via_repo(monkeypatch, tmp_plita) -> None:
    load_result = load(
        PlanBuildInput(
            start_date="2026-04-21",
            tracks_count=3,
            filter_method="all",
        ),
        config=LoadConfig(plita_db_path=tmp_plita, pb_db_path=tmp_plita),
        plan_load=_plan_load_for_db(tmp_plita),
    )
    order = load_result.orders_2d[0]
    tracks = [
        {
            "label": "ОСНОВНАЯ",
            "items": [
                {
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"] / 1000.0,
                    "load_code": order["load_code"],
                }
            ],
        }
    ]
    from core.production.dto import OptimizeResult

    opt_result = OptimizeResult(
        all_tracks_list=tracks,
        optimization_result={
            "total_plates": 1,
            "plate_assignments": [
                {
                    "source": "primary",
                    "kp_id": order["kp_id"],
                    "plate_name": order["plate_name"],
                    "length": order["length"],
                    "width": order["width"],
                    "load_code": order["load_code"],
                }
            ],
        },
    )

    monkeypatch.setattr(
        "core.production.planning.commit_plan_plates",
        lambda **kwargs: None,
    )

    repo = _FakePersistRepo()
    result = persist(
        load_result,
        opt_result,
        PersistConfig(
            plita_db_path=tmp_plita,
            start_date="2026-04-21",
            tracks_count=3,
        ),
        repo,
        ensure_unique_plan_id=lambda: "plan_unique",
    )

    assert result.plan["id"] in repo.plans
    assert repo.active_id == result.plan["id"]
    assert result.summary["total_tracks"] == 1
