"""SGP service unit tests: unlink/relink/split/qty balance."""

from __future__ import annotations

import sqlite3

import pytest

from app.domain.enums import KpStatus, PlateTransitionReason
from app.services.sgp_service import SgpError, SgpService
from core import kp_db
from tests.helpers import kp_db_fixtures as fx


PLATE = "ПБ 60-12-8п"


@pytest.fixture
def db(tmp_path) -> str:
    return fx.make_iso_db(tmp_path)


def _seed_linked_sgp(db_path: str, *, kp_id: int = 1, qty: int = 5) -> int:
    fx.seed_kp_offer(db_path, kp_id, customer_name="КлиентА")
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET ordered_qty = ? WHERE kp_id = ?",
            (qty, kp_id),
        )
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (?, ?, 6.0, 1.2, 800, ?, '27.07.2026', 1, 'plan-1')
            """,
            (kp_id, PLATE, qty),
        )
        conn.commit()
        return int(cur.lastrowid)


def test_list_plates_filters(db: str) -> None:
    sgp_id = _seed_linked_sgp(db, qty=3)
    svc = SgpService(db_path=db)
    with sqlite3.connect(db) as conn:
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date
            ) VALUES (NULL, ?, 6.0, 1.2, 800, 2, '27.07.2026')
            """,
            (PLATE,),
        )
        conn.commit()

    assert svc.list_plates(filter="all").count == 2
    assert svc.list_plates(filter="linked").count == 1
    assert svc.list_plates(filter="unlinked").count == 1
    linked = svc.list_plates(filter="linked").items[0]
    assert linked.id == sgp_id
    assert linked.sgp_progress is not None
    assert linked.sgp_progress.n == 3
    assert linked.sgp_progress.m == 3


def test_unlink_partial_split_and_balance(db: str) -> None:
    sgp_id = _seed_linked_sgp(db, qty=5)
    svc = SgpService(db_path=db)

    before_kp = 0
    before_sgp = 5
    svc.unlink(sgp_id, 2, actor="test")

    with sqlite3.connect(db) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COALESCE(SUM(qty),0) FROM kp_plates WHERE kp_id=1")
        after_kp = int(cur.fetchone()[0])
        cur.execute("SELECT COALESCE(SUM(qty),0) FROM completed_plates")
        after_sgp = int(cur.fetchone()[0])
        cur.execute(
            "SELECT qty, kp_id FROM completed_plates ORDER BY id"
        )
        rows = cur.fetchall()
        cur.execute(
            "SELECT reason, qty FROM plate_status_log WHERE reason = ?",
            (PlateTransitionReason.SGP_UNLINK.value,),
        )
        audit = cur.fetchone()
        cur.execute("SELECT status FROM kp_meta WHERE kp_id=1")
        status = cur.fetchone()[0]

    assert after_kp == before_kp + 2
    assert after_sgp == before_sgp  # physics unchanged
    assert sorted((r[0], r[1]) for r in rows) == [(2, None), (3, 1)] or sorted(
        (r[0], r[1]) for r in rows
    ) == [(3, 1), (2, None)]
    assert { (r[0], r[1]) for r in rows } == {(3, 1), (2, None)}
    assert audit == (PlateTransitionReason.SGP_UNLINK.value, 2)
    assert status == KpStatus.IN_WORK.value


def test_relink_strict_match_and_reject(db: str) -> None:
    sgp_id = _seed_linked_sgp(db, kp_id=1, qty=5)
    fx.seed_kp_offer(db, 2, customer_name="КлиентБ")
    fx.seed_plate(
        db,
        kp_id=2,
        plate_name=PLATE,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
        qty=4,
        status="в производстве",
    )
    # Wrong dims demand for KP 3
    fx.seed_kp_offer(db, 3, customer_name="КлиентВ")
    fx.seed_plate(
        db,
        kp_id=3,
        plate_name=PLATE,
        length_m=5.0,
        width_m=1.2,
        load_class=800,
        qty=4,
        status="в производстве",
    )

    svc = SgpService(db_path=db)
    svc.unlink(sgp_id, 5)
    free = svc.list_plates(filter="unlinked").items[0]

    with pytest.raises(SgpError) as exc:
        svc.relink(free.id, target_kp_id=3, qty=2)
    assert exc.value.code == "sgp_no_matching_demand"

    with sqlite3.connect(db) as conn:
        snap = conn.execute(
            "SELECT COALESCE(SUM(qty),0) FROM completed_plates WHERE kp_id IS NULL"
        ).fetchone()[0]

    result = svc.relink(free.id, target_kp_id=2, qty=2)
    assert result.ok
    assert result.target_kp_id == 2

    with sqlite3.connect(db) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT COALESCE(SUM(qty),0) FROM kp_plates WHERE kp_id=2"
        )
        assert int(cur.fetchone()[0]) == 2  # 4-2
        cur.execute(
            "SELECT COALESCE(SUM(qty),0) FROM completed_plates WHERE kp_id=2"
        )
        assert int(cur.fetchone()[0]) == 2
        cur.execute(
            "SELECT COALESCE(SUM(qty),0) FROM completed_plates WHERE kp_id IS NULL"
        )
        assert int(cur.fetchone()[0]) == 3  # 5-2 remaining free
        # Reject path left DB unchanged for wrong target
        assert int(snap) == 5


def test_sgp_progress_n_m(db: str) -> None:
    _seed_linked_sgp(db, qty=10)
    svc = SgpService(db_path=db)
    progress = svc.sgp_progress(1)
    assert progress.n == 10
    assert progress.m == 10

    linked = svc.list_plates(filter="linked").items[0]
    svc.unlink(linked.id, 2)
    progress2 = svc.sgp_progress(1)
    assert progress2.n == 8
    assert progress2.m == 10


def test_sgp_progress_read_does_not_persist_ordered_qty(db: str) -> None:
    """Read-path must not freeze/UPDATE ordered_qty (audit Q4/A9)."""
    fx.seed_kp_offer(db, 1, customer_name="КлиентА")
    fx.seed_plate(
        db,
        kp_id=1,
        plate_name=PLATE,
        length_m=6.0,
        width_m=1.2,
        qty=4,
        status="в производстве",
    )
    with sqlite3.connect(db) as conn:
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (1, ?, 6.0, 1.2, 800, 2, '27.07.2026', 1, 'plan-1')
            """,
            (PLATE,),
        )
        conn.commit()
        before = conn.execute(
            "SELECT ordered_qty FROM kp_meta WHERE kp_id = 1"
        ).fetchone()[0]

    assert before is None
    svc = SgpService(db_path=db)
    progress = svc.sgp_progress(1)
    assert progress.n == 2
    assert progress.m == 6  # ephemeral: remaining 4 + on_sgp 2

    with sqlite3.connect(db) as conn:
        after = conn.execute(
            "SELECT ordered_qty FROM kp_meta WHERE kp_id = 1"
        ).fetchone()[0]
    assert after is None
