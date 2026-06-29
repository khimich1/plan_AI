"""Migrate production plans from JSON files to SQLite (WP4 / A2 part 2).

Reads ``bot/data/plans/*.json`` and ``bot/data/plans_metadata.json``, writes
rows into ``production_plans`` via :class:`app.repositories.plan_repository.PlanRepository`.

Features:
- Idempotent: existing plans with matching payload are skipped.
- ``--dry-run``: report only, no DB writes and no backup.
- Real run: timestamped backup of source files before migration.
- Post-migration verification: stored payload matches normalized source.

Usage::

    python scripts/migrate_plans_to_sqlite.py [--dry-run]
    python scripts/migrate_plans_to_sqlite.py --plans-dir PATH --metadata-path PATH --db-path PATH
"""
from __future__ import annotations

import argparse
import json
import logging
import shutil
import sqlite3
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.repositories.plan_repository import PlanRepository  # noqa: E402
from core.serialization import strip_plate_audit_from_plan  # noqa: E402

logger = logging.getLogger("migrate_plans_to_sqlite")

DEFAULT_PLANS_DIR = PROJECT_ROOT / "bot" / "data" / "plans"
DEFAULT_METADATA_PATH = PROJECT_ROOT / "bot" / "data" / "plans_metadata.json"
DEFAULT_DB_PATH = PROJECT_ROOT / "plita.db"

PLAN_KEY_FIELDS = ("id", "name", "created_at", "start_date", "tracks_count")


@dataclass
class PlanMigrationOutcome:
    plan_id: str
    status: str
    message: str = ""


@dataclass
class MigrationReport:
    created: int = 0
    skipped: int = 0
    errors: int = 0
    verified: int = 0
    active_plan_id: str | None = None
    backup_dir: Path | None = None
    outcomes: list[PlanMigrationOutcome] = field(default_factory=list)

    def add(self, outcome: PlanMigrationOutcome) -> None:
        self.outcomes.append(outcome)
        if outcome.status == "created":
            self.created += 1
        elif outcome.status == "skipped":
            self.skipped += 1
        elif outcome.status == "dry_run":
            self.created += 1
        elif outcome.status == "error":
            self.errors += 1
        elif outcome.status == "verified":
            self.verified += 1


def normalize_plan_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """Apply the same normalization used when persisting to SQLite."""
    return strip_plate_audit_from_plan(payload)


def load_json_file(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"expected JSON object in {path}")
    return data


def load_plans_metadata(metadata_path: Path) -> dict[str, Any]:
    if not metadata_path.exists():
        return {"plans": [], "active_plan_id": None}
    return load_json_file(metadata_path)


def discover_plan_files(plans_dir: Path) -> list[Path]:
    if not plans_dir.exists():
        return []
    return sorted(plans_dir.glob("*.json"))


def load_source_plans(plans_dir: Path) -> tuple[list[dict[str, Any]], list[str]]:
    """Load plan payloads from JSON files. Returns (plans, load_errors)."""
    plans: list[dict[str, Any]] = []
    errors: list[str] = []

    for plan_path in discover_plan_files(plans_dir):
        try:
            payload = load_json_file(plan_path)
        except Exception as exc:
            errors.append(f"{plan_path.name}: failed to read JSON ({exc})")
            continue

        plan_id = payload.get("id")
        if not plan_id or not isinstance(plan_id, str):
            errors.append(f"{plan_path.name}: missing or invalid 'id'")
            continue

        if plan_path.stem != plan_id:
            logger.warning(
                "Filename %s does not match plan id %r; using payload id",
                plan_path.name,
                plan_id,
            )

        plans.append(payload)

    return plans, errors


def compare_plan_payloads(source: dict[str, Any], stored: dict[str, Any]) -> list[str]:
    """Compare normalized source with stored payload; return mismatch messages."""
    normalized = normalize_plan_payload(source)
    mismatches: list[str] = []

    for key in PLAN_KEY_FIELDS:
        if normalized.get(key) != stored.get(key):
            mismatches.append(
                f"field '{key}': source={normalized.get(key)!r}, stored={stored.get(key)!r}"
            )

    if normalized.get("days") != stored.get("days"):
        src_days = normalized.get("days") or {}
        stored_days = stored.get("days") or {}
        mismatches.append(
            f"days mismatch: source_days={len(src_days)}, stored_days={len(stored_days)}"
        )

    return mismatches


def verify_plan_in_db(
    repository: PlanRepository,
    source_payload: dict[str, Any],
) -> list[str]:
    plan_id = source_payload["id"]
    loaded = repository.get(plan_id)
    if loaded is None:
        return [f"plan {plan_id} not found in database after migration"]
    return compare_plan_payloads(source_payload, loaded["payload"])


def backup_source_files(
    *,
    plans_dir: Path,
    metadata_path: Path,
    backup_root: Path,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = backup_root / f"plans_migration_{timestamp}"
    backup_dir.mkdir(parents=True, exist_ok=True)

    if plans_dir.exists():
        shutil.copytree(plans_dir, backup_dir / "plans", dirs_exist_ok=True)

    if metadata_path.exists():
        shutil.copy2(metadata_path, backup_dir / metadata_path.name)

    return backup_dir


def migrate_plan(
    repository: PlanRepository | None,
    source_payload: dict[str, Any],
    *,
    dry_run: bool,
) -> PlanMigrationOutcome:
    plan_id = source_payload["id"]

    if dry_run and repository is None:
        return PlanMigrationOutcome(plan_id, "dry_run", "would create")

    if repository is None:
        return PlanMigrationOutcome(plan_id, "error", "database repository is not available")

    existing = repository.get(plan_id)
    if existing is not None:
        mismatches = compare_plan_payloads(source_payload, existing["payload"])
        if not mismatches:
            return PlanMigrationOutcome(plan_id, "skipped", "already in database")
        return PlanMigrationOutcome(
            plan_id,
            "error",
            "plan exists with different payload: " + "; ".join(mismatches),
        )

    if dry_run:
        return PlanMigrationOutcome(plan_id, "dry_run", "would create")

    try:
        created = repository.create(source_payload)
    except sqlite3.IntegrityError as exc:
        return PlanMigrationOutcome(plan_id, "error", f"database insert failed: {exc}")

    if created["version"] != 1:
        return PlanMigrationOutcome(
            plan_id,
            "error",
            f"unexpected version after create: {created['version']}",
        )

    mismatches = compare_plan_payloads(source_payload, created["payload"])
    if mismatches:
        return PlanMigrationOutcome(
            plan_id,
            "error",
            "verification failed after create: " + "; ".join(mismatches),
        )

    return PlanMigrationOutcome(plan_id, "created")


def apply_active_plan(
    repository: PlanRepository | None,
    active_plan_id: str | None,
    *,
    dry_run: bool,
) -> str | None:
    if not active_plan_id:
        return None

    if dry_run:
        logger.info("[dry-run] would set active plan: %s", active_plan_id)
        return active_plan_id

    if repository is None:
        logger.warning("Cannot set active plan: database repository is not available")
        return None

    if not repository.set_active(active_plan_id):
        logger.warning("active_plan_id %r not found in database", active_plan_id)
        return None

    return active_plan_id


def run_migration(
    *,
    plans_dir: Path,
    metadata_path: Path,
    db_path: Path,
    dry_run: bool = False,
    backup_root: Path | None = None,
) -> MigrationReport:
    report = MigrationReport()
    repository: PlanRepository | None
    if db_path.exists() or not dry_run:
        repository = PlanRepository(str(db_path))
    else:
        repository = None

    source_plans, load_errors = load_source_plans(plans_dir)
    for message in load_errors:
        report.add(PlanMigrationOutcome("?", "error", message))
        logger.error(message)

    metadata = load_plans_metadata(metadata_path)
    active_plan_id = metadata.get("active_plan_id")

    if not dry_run and (source_plans or metadata_path.exists()):
        root = backup_root or metadata_path.parent
        report.backup_dir = backup_source_files(
            plans_dir=plans_dir,
            metadata_path=metadata_path,
            backup_root=root,
        )
        logger.info("Backup created at %s", report.backup_dir)

    for payload in source_plans:
        outcome = migrate_plan(repository, payload, dry_run=dry_run)
        report.add(outcome)
        if outcome.status == "error":
            logger.error("Plan %s: %s", outcome.plan_id, outcome.message)
        elif outcome.status == "skipped":
            logger.info("Plan %s: skipped (%s)", outcome.plan_id, outcome.message)
        elif outcome.status == "dry_run":
            logger.info("Plan %s: dry-run (%s)", outcome.plan_id, outcome.message)
        else:
            logger.info("Plan %s: migrated", outcome.plan_id)

    report.active_plan_id = apply_active_plan(
        repository,
        active_plan_id if isinstance(active_plan_id, str) else None,
        dry_run=dry_run,
    )

    for payload in source_plans:
        plan_id = payload["id"]
        if dry_run:
            continue
        if any(
            o.plan_id == plan_id and o.status in {"created", "skipped"}
            for o in report.outcomes
        ):
            mismatches = verify_plan_in_db(repository, payload)
            if mismatches:
                report.add(
                    PlanMigrationOutcome(
                        plan_id,
                        "error",
                        "post-migration verification failed: " + "; ".join(mismatches),
                    )
                )
            else:
                report.verified += 1

    return report


def print_report(report: MigrationReport, *, dry_run: bool) -> None:
    mode = "DRY-RUN" if dry_run else "MIGRATION"
    print(f"\n=== Plans {mode} report ===")
    print(f"Created:   {report.created}")
    print(f"Skipped:   {report.skipped}")
    print(f"Verified:  {report.verified}")
    print(f"Errors:    {report.errors}")
    if report.active_plan_id:
        print(f"Active:    {report.active_plan_id}")
    if report.backup_dir:
        print(f"Backup:    {report.backup_dir}")

    if report.errors:
        print("\nErrors:")
        for outcome in report.outcomes:
            if outcome.status == "error":
                print(f"  - {outcome.plan_id}: {outcome.message}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--plans-dir",
        default=str(DEFAULT_PLANS_DIR),
        help=f"Directory with plan JSON files (default: {DEFAULT_PLANS_DIR})",
    )
    parser.add_argument(
        "--metadata-path",
        default=str(DEFAULT_METADATA_PATH),
        help=f"Path to plans_metadata.json (default: {DEFAULT_METADATA_PATH})",
    )
    parser.add_argument(
        "--db-path",
        default=str(DEFAULT_DB_PATH),
        help=f"SQLite database path (default: {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--backup-root",
        default=None,
        help="Directory for migration backups (default: parent of metadata file)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report actions without writing to DB or creating backups",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="[%(asctime)s] %(levelname)s %(message)s",
    )

    args = build_parser().parse_args(argv)
    plans_dir = Path(args.plans_dir)
    metadata_path = Path(args.metadata_path)
    db_path = Path(args.db_path)
    backup_root = Path(args.backup_root) if args.backup_root else None

    if not db_path.exists() and not args.dry_run:
        logger.error("Database not found: %s", db_path)
        return 1

    if not plans_dir.exists():
        logger.warning("Plans directory not found: %s (nothing to migrate)", plans_dir)

    report = run_migration(
        plans_dir=plans_dir,
        metadata_path=metadata_path,
        db_path=db_path,
        dry_run=args.dry_run,
        backup_root=backup_root,
    )
    print_report(report, dry_run=args.dry_run)

    return 1 if report.errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
