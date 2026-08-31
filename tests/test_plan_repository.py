from __future__ import annotations

from pathlib import Path

import pytest

from app.repositories.plan_errors import PlanVersionConflict
from app.repositories.plan_repository import PlanRepository


def _sample_payload(plan_id: str = "plan_test_001", *, name: str = "Test plan") -> dict:
    return {
        "id": plan_id,
        "name": name,
        "created_at": "2026-06-19 10:00:00",
        "start_date": "2026-06-20",
        "tracks_count": 5,
        "days": {
            "1": {
                "tracks": [{"track_number": 1, "plates": []}],
                "saved_tracks_count": 1,
            }
        },
    }


def make_repository(tmp_path: Path) -> PlanRepository:
    return PlanRepository(str(tmp_path / "plans.db"))


def test_create_and_get_round_trip(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    payload = _sample_payload()

    created = repository.create(payload)
    loaded = repository.get("plan_test_001")

    assert created["version"] == 1
    assert created["payload"]["id"] == "plan_test_001"
    assert loaded is not None
    assert loaded["version"] == 1
    assert loaded["payload"]["name"] == "Test plan"
    assert loaded["payload"]["days"]["1"]["tracks"][0]["track_number"] == 1


def test_save_increments_version(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    payload = _sample_payload()
    repository.create(payload)

    payload["name"] = "Updated plan"
    saved = repository.save(payload, expected_version=1)
    loaded = repository.get("plan_test_001")

    assert saved["version"] == 2
    assert saved["payload"]["name"] == "Updated plan"
    assert loaded is not None
    assert loaded["version"] == 2
    assert loaded["payload"]["name"] == "Updated plan"


def test_save_with_stale_version_raises_conflict_and_preserves_data(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    payload = _sample_payload()
    repository.create(payload)

    stale_payload = dict(payload)
    stale_payload["name"] = "Stale write"
    repository.save(dict(payload, name="Current write"), expected_version=1)

    with pytest.raises(PlanVersionConflict) as exc_info:
        repository.save(stale_payload, expected_version=1)

    conflict = exc_info.value
    assert conflict.plan_id == "plan_test_001"
    assert conflict.expected_version == 1

    loaded = repository.get("plan_test_001")
    assert loaded is not None
    assert loaded["version"] == 2
    assert loaded["payload"]["name"] == "Current write"


def test_delete_removes_plan(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    repository.create(_sample_payload())

    assert repository.delete("plan_test_001") is True
    assert repository.get("plan_test_001") is None
    assert repository.delete("plan_test_001") is False


def test_list_metadata_from_sqlite(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    repository.create(_sample_payload("plan_a", name="Plan A"))
    repository.create(_sample_payload("plan_b", name="Plan B"))
    repository.set_active("plan_b")

    metadata = repository.list_metadata()

    assert metadata["active_plan_id"] == "plan_b"
    assert len(metadata["plans"]) == 2
    ids = {entry["id"] for entry in metadata["plans"]}
    assert ids == {"plan_a", "plan_b"}
    plan_b_meta = next(entry for entry in metadata["plans"] if entry["id"] == "plan_b")
    assert plan_b_meta["name"] == "Plan B"
    assert plan_b_meta["total_days"] == 1
    assert plan_b_meta["total_tracks"] == 1


def test_get_active_and_set_active(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    repository.create(_sample_payload("plan_a"))
    repository.create(_sample_payload("plan_b"))

    assert repository.get_active() is None
    assert repository.set_active("plan_b") is True

    active = repository.get_active()
    assert active is not None
    assert active["payload"]["id"] == "plan_b"
    assert active["version"] == 1
    assert repository.get_active_plan_id() == "plan_b"

    assert repository.set_active("missing") is False


def test_create_requires_plan_id(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)

    with pytest.raises(ValueError, match="'id'"):
        repository.create({"name": "no id"})


def test_concurrent_stale_writes_only_first_succeeds(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    base = _sample_payload()
    repository.create(base)

    first_writer = dict(base, name="Writer 1")
    second_writer = dict(base, name="Writer 2")

    repository.save(first_writer, expected_version=1)

    with pytest.raises(PlanVersionConflict):
        repository.save(second_writer, expected_version=1)

    loaded = repository.get("plan_test_001")
    assert loaded is not None
    assert loaded["payload"]["name"] == "Writer 1"
    assert loaded["version"] == 2


def test_mark_day_completed_on_external_conn_rollback_does_not_persist(
    tmp_path: Path,
) -> None:
    repository = make_repository(tmp_path)
    repository.create(_sample_payload())

    conn = repository._connect()
    try:
        assert (
            repository.mark_day_completed(
                "plan_test_001",
                "1",
                _external_conn=conn,
            )
            is True
        )
        conn.rollback()
    finally:
        conn.close()

    loaded = repository.get("plan_test_001")
    assert loaded is not None
    assert loaded["version"] == 1
    assert loaded["payload"]["days"]["1"].get("completed") is not True
    assert loaded["payload"].get("completed_days", []) == []
