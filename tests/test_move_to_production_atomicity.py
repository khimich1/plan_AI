"""MTP-110/300: atomic move-to-production (execution_terms + status + freeze M)."""

from __future__ import annotations

import sqlite3

import pytest

from core.kp import offers_write
from tests.helpers.kp_db_fixtures import make_iso_db, seed_kp_offer, seed_plate


def _state(db_path: str, kp_id: int) -> dict:
    with sqlite3.connect(db_path) as conn:
        meta = conn.execute(
            "SELECT status, ordered_qty FROM kp_meta WHERE kp_id = ?", (kp_id,)
        ).fetchone()
        terms = conn.execute(
            "SELECT execution_terms FROM KP_offers WHERE kp_id = ?", (kp_id,)
        ).fetchone()
    return {
        "status": meta[0] if meta else None,
        "ordered_qty": meta[1] if meta else None,
        "execution_terms": terms[0] if terms else None,
    }


def test_commit_move_to_production_happy_path(tmp_path) -> None:
    db_path = make_iso_db(tmp_path)
    seed_kp_offer(db_path, 10, status="в архиве")
    seed_plate(
        db_path,
        kp_id=10,
        plate_name="ПБ 60-12-8",
        length_m=6.0,
        width_m=1.2,
        qty=5,
        status="в производстве",
    )

    ordered = offers_write.commit_move_to_production(10, "15.08.2026", db_path)

    assert ordered == 5
    st = _state(db_path, 10)
    assert st["status"] == "в работе"
    assert st["execution_terms"] == "15.08.2026"
    assert st["ordered_qty"] == 5


def test_commit_move_to_production_rolls_back_when_freeze_raises(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = make_iso_db(tmp_path)
    seed_kp_offer(db_path, 11, status="в архиве")
    seed_plate(
        db_path,
        kp_id=11,
        plate_name="ПБ 60-12-8",
        length_m=6.0,
        width_m=1.2,
        qty=3,
        status="в производстве",
    )
    before = _state(db_path, 11)

    def boom(cur, kp_id):
        raise RuntimeError("freeze failed")

    monkeypatch.setattr(
        "core.kp_db_plates_completion.freeze_ordered_qty_if_needed",
        boom,
    )

    with pytest.raises(RuntimeError, match="freeze failed"):
        offers_write.commit_move_to_production(11, "20.08.2026", db_path)

    after = _state(db_path, 11)
    assert after["status"] == before["status"] == "в архиве"
    assert after["execution_terms"] == before["execution_terms"]
    assert after["ordered_qty"] is None


def test_commit_move_to_production_fails_when_freeze_returns_none(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = make_iso_db(tmp_path)
    seed_kp_offer(db_path, 12, status="в архиве")

    monkeypatch.setattr(
        "core.kp_db_plates_completion.freeze_ordered_qty_if_needed",
        lambda cur, kp_id: None,
    )

    with pytest.raises(ValueError, match="freeze_ordered_qty_failed"):
        offers_write.commit_move_to_production(12, "20.08.2026", db_path)

    st = _state(db_path, 12)
    assert st["status"] == "в архиве"
    assert st["ordered_qty"] is None


def _seed_archived_with_plates(db_path: str, kp_id: int, qty: int = 4) -> None:
    seed_kp_offer(db_path, kp_id, status="в архиве")
    seed_plate(
        db_path,
        kp_id=kp_id,
        plate_name="ПБ 60-12-8",
        length_m=6.0,
        width_m=1.2,
        qty=qty,
        status="в производстве",
    )


def test_archive_service_move_to_production_freezes_m(tmp_path) -> None:
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveService

    db_path = make_iso_db(tmp_path)
    _seed_archived_with_plates(db_path, 20, qty=4)
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "out",
    )
    admin = {"id": 1, "role": "admin"}

    details = service.move_to_production(20, "15.08.2026", user=admin)

    assert details.status == "в работе"
    assert details.execution_terms == "15.08.2026"
    assert _state(db_path, 20)["ordered_qty"] == 4


def test_archive_service_move_to_production_rolls_back_on_freeze_fail(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveError, ArchiveService

    db_path = make_iso_db(tmp_path)
    _seed_archived_with_plates(db_path, 21, qty=2)
    before = _state(db_path, 21)
    monkeypatch.setattr(
        "core.kp_db_plates_completion.freeze_ordered_qty_if_needed",
        lambda cur, kp_id: (_ for _ in ()).throw(RuntimeError("freeze boom")),
    )
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "out",
    )

    with pytest.raises(ArchiveError):
        service.move_to_production(21, "15.08.2026", user={"id": 1, "role": "admin"})

    after = _state(db_path, 21)
    assert after["status"] == before["status"] == "в архиве"
    assert after["execution_terms"] == before["execution_terms"]
    assert after["ordered_qty"] is None


def test_offers_service_move_to_production_freezes_m(tmp_path) -> None:
    from app.repositories.kp_repository import KpRepository
    from app.services.offers_service import OffersService

    db_path = make_iso_db(tmp_path)
    _seed_archived_with_plates(db_path, 30, qty=7)
    service = OffersService(kp_repository=KpRepository(db_path=db_path))
    admin = {"id": 1, "role": "admin"}

    result = service.move_to_production(30, "15.08.2026", user=admin)

    assert result["execution_terms"] == "15.08.2026"
    assert result["offer"]["status"] == "в работе"
    assert _state(db_path, 30)["ordered_qty"] == 7


def test_offers_service_move_to_production_rolls_back_on_freeze_fail(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.repositories.kp_repository import KpRepository
    from app.services.offers_service import OffersService

    db_path = make_iso_db(tmp_path)
    _seed_archived_with_plates(db_path, 31, qty=2)
    before = _state(db_path, 31)
    monkeypatch.setattr(
        "core.kp_db_plates_completion.freeze_ordered_qty_if_needed",
        lambda cur, kp_id: (_ for _ in ()).throw(RuntimeError("freeze boom")),
    )
    service = OffersService(kp_repository=KpRepository(db_path=db_path))

    with pytest.raises(ValueError, match="move_to_production_failed"):
        service.move_to_production(31, "15.08.2026", user={"id": 1, "role": "admin"})

    after = _state(db_path, 31)
    assert after["status"] == before["status"] == "в архиве"
    assert after["ordered_qty"] is None
