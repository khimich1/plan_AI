from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.repositories.plan_repository import PlanRepository
from scripts.migrate_plans_to_sqlite import (
    compare_plan_payloads,
    main,
    run_migration,
)


def _sample_plan(
    plan_id: str,
    *,
    name: str = "Fixture plan",
) -> dict:
    return {
        "id": plan_id,
        "name": name,
        "created_at": "2026-06-19 10:00:00",
        "start_date": "2026-06-20",
        "tracks_count": 5,
        "days": {
            "2026-06-20": {
                "date": "2026-06-20",
                "day_number": 1,
                "tracks": [{"track_number": 1, "items": []}],
                "saved_tracks_count": 1,
            }
        },
    }


@pytest.fixture
def fixture_plans_dir(tmp_path: Path) -> tuple[Path, Path, Path]:
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    metadata_path = tmp_path / "plans_metadata.json"
    db_path = tmp_path / "plita.db"

    plan_a = _sample_plan("plan_a", name="Plan A")
    plan_b = _sample_plan("plan_b", name="Plan B")

    (plans_dir / "plan_a.json").write_text(
        json.dumps(plan_a, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (plans_dir / "plan_b.json").write_text(
        json.dumps(plan_b, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    metadata_path.write_text(
        json.dumps(
            {
                "plans": [
                    {"id": "plan_a", "name": "Plan A"},
                    {"id": "plan_b", "name": "Plan B"},
                ],
                "active_plan_id": "plan_b",
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return plans_dir, metadata_path, db_path


def test_compare_plan_payloads_matches_normalized_source() -> None:
    source = _sample_plan("plan_x")
    stored = dict(source)
    assert compare_plan_payloads(source, stored) == []


def test_dry_run_does_not_write_db_or_backup(
    fixture_plans_dir: tuple[Path, Path, Path],
) -> None:
    plans_dir, metadata_path, db_path = fixture_plans_dir

    report = run_migration(
        plans_dir=plans_dir,
        metadata_path=metadata_path,
        db_path=db_path,
        dry_run=True,
    )

    assert report.created == 2
    assert report.errors == 0
    assert report.backup_dir is None
    assert not db_path.exists()
    assert report.active_plan_id == "plan_b"


def test_migration_creates_plans_and_sets_active(
    fixture_plans_dir: tuple[Path, Path, Path],
) -> None:
    plans_dir, metadata_path, db_path = fixture_plans_dir

    report = run_migration(
        plans_dir=plans_dir,
        metadata_path=metadata_path,
        db_path=db_path,
        dry_run=False,
        backup_root=metadata_path.parent / "backups",
    )

    assert report.errors == 0
    assert report.created == 2
    assert report.verified == 2
    assert report.active_plan_id == "plan_b"
    assert report.backup_dir is not None
    assert (report.backup_dir / "plans" / "plan_a.json").exists()
    assert (report.backup_dir / "plans_metadata.json").exists()

    repository = PlanRepository(str(db_path))
    plan_a = repository.get("plan_a")
    plan_b = repository.get("plan_b")
    active = repository.get_active()

    assert plan_a is not None
    assert plan_b is not None
    assert plan_a["version"] == 1
    assert plan_b["version"] == 1
    assert active is not None
    assert active["payload"]["id"] == "plan_b"
    assert repository.get_active_plan_id() == "plan_b"


def test_migration_is_idempotent(
    fixture_plans_dir: tuple[Path, Path, Path],
) -> None:
    plans_dir, metadata_path, db_path = fixture_plans_dir

    first = run_migration(
        plans_dir=plans_dir,
        metadata_path=metadata_path,
        db_path=db_path,
        dry_run=False,
        backup_root=metadata_path.parent / "backups",
    )
    second = run_migration(
        plans_dir=plans_dir,
        metadata_path=metadata_path,
        db_path=db_path,
        dry_run=False,
        backup_root=metadata_path.parent / "backups",
    )

    assert first.errors == 0
    assert first.created == 2
    assert second.errors == 0
    assert second.created == 0
    assert second.skipped == 2
    assert second.verified == 2


def test_main_dry_run_exits_zero_without_db(
    fixture_plans_dir: tuple[Path, Path, Path],
) -> None:
    plans_dir, metadata_path, db_path = fixture_plans_dir
    missing_db = db_path.parent / "missing.db"

    exit_code = main(
        [
            "--dry-run",
            "--plans-dir",
            str(plans_dir),
            "--metadata-path",
            str(metadata_path),
            "--db-path",
            str(missing_db),
        ]
    )

    assert exit_code == 0
    assert not missing_db.exists()
