#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Разовая заливка однозначных GUID не-плит в nomenclature_guid (GPS-003).

Берёт ✅-совпадения того же матчера, что собирает Excel-очередь
(scripts/build_guid_queue.py): только plain GUID, без ⚠️/🏭, без плит.
Карточка «у» пишется в guid_1c_u той же марки. Caller коммитит.

Запуск из корня репозитория:

    .venv/bin/python scripts/bootstrap_nomenclature_guid.py --dry-run
    .venv/bin/python scripts/bootstrap_nomenclature_guid.py
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.nomenclature_guid import ensure_schema, get_by_mark, upsert
from core.project_paths import PRICE_DB_PATH
from scripts.build_guid_queue import (
    DEFAULT_GUID_DIR,
    NON_PLATE_QUEUE_GROUPS,
    PRICE_FILE_NAMES,
    QUEUE_GROUP_TO_KIND,
    iter_our_catalog,
    load_1c,
    load_db,
)

MATCH_STATUS_AUTO = "auto"
MATCH_STATUS_MANUAL = "manual"
KIND_ORDER = ("pile", "bridge_pile", "fbs", "stair_flight", "stair_step")


@dataclass(frozen=True, slots=True)
class BootstrapCandidate:
    product_kind: str
    mark: str
    guid_1c: str
    guid_1c_u: Optional[str] = None


@dataclass
class BootstrapReport:
    inserted: int = 0
    updated: int = 0
    skipped_manual: int = 0
    skipped_dup: int = 0
    skipped_not_ready: int = 0
    by_kind: dict[str, int] = field(default_factory=dict)
    with_guid_u: int = 0
    skipped_dup_marks: tuple[str, ...] = ()


def _guid_eq(left: Optional[str], right: Optional[str]) -> bool:
    return (left or "").strip().lower() == (right or "").strip().lower()


def _row_unchanged(
    existing_guid_1c: Optional[str],
    existing_guid_1c_u: Optional[str],
    existing_status: str,
    candidate: BootstrapCandidate,
) -> bool:
    return (
        existing_status == MATCH_STATUS_AUTO
        and _guid_eq(existing_guid_1c, candidate.guid_1c)
        and _guid_eq(existing_guid_1c_u, candidate.guid_1c_u)
    )


def collect_auto_candidates(
    name_idx,
    db: dict,
    *,
    groups: Iterable[str] = NON_PLATE_QUEUE_GROUPS,
) -> tuple[list[BootstrapCandidate], int, int, tuple[str, ...]]:
    """Ready unique-GUID matches. Returns (candidates, skipped_not_ready, skipped_dup, dup labels)."""
    skipped_not_ready = 0
    group_list = tuple(groups)
    if "plates" in group_list:
        skipped_not_ready += len(db.get("plates") or [])
    safe_groups = tuple(
        key for key in group_list if key != "plates" and key in QUEUE_GROUP_TO_KIND
    )
    ready: list[BootstrapCandidate] = []
    for group_key, _db_key, db_label, _src, matched in iter_our_catalog(
        name_idx, db, groups=safe_groups
    ):
        kind = QUEUE_GROUP_TO_KIND[group_key]
        if matched.status != "ok" or matched.item is None:
            skipped_not_ready += 1
            continue
        ready.append(
            BootstrapCandidate(
                product_kind=kind,
                mark=db_label,
                guid_1c=matched.item.guid,
                guid_1c_u=matched.u_item.guid if matched.u_item else None,
            )
        )

    collapsed: dict[tuple[str, str], BootstrapCandidate] = {}
    dup_keys: set[tuple[str, str]] = set()
    for cand in ready:
        key = (cand.product_kind, cand.mark)
        prev = collapsed.get(key)
        if prev is None:
            collapsed[key] = cand
            continue
        if prev != cand:
            dup_keys.add(key)

    guid_owners: dict[str, set[tuple[str, str]]] = defaultdict(set)
    for key, cand in collapsed.items():
        if key in dup_keys:
            continue
        guid_owners[cand.guid_1c].add(key)
        if cand.guid_1c_u:
            guid_owners[cand.guid_1c_u].add(key)
    for owners in guid_owners.values():
        if len(owners) > 1:
            dup_keys.update(owners)

    unique: list[BootstrapCandidate] = []
    dup_labels = tuple(
        sorted(f"{kind}:{mark}" for kind, mark in dup_keys)
    )
    for key, cand in collapsed.items():
        if key in dup_keys:
            continue
        unique.append(cand)
    return unique, skipped_not_ready, len(dup_labels), dup_labels


def apply_bootstrap(
    conn: sqlite3.Connection,
    candidates: Iterable[BootstrapCandidate],
    *,
    dry_run: bool = False,
) -> BootstrapReport:
    """Insert/update auto rows. Does not commit. Does not clobber manual."""
    table_ready = True
    if dry_run:
        table_ready = (
            conn.execute(
                "SELECT 1 FROM sqlite_master "
                "WHERE type='table' AND name='nomenclature_guid'"
            ).fetchone()
            is not None
        )
    else:
        ensure_schema(conn)
    report = BootstrapReport(by_kind={kind: 0 for kind in KIND_ORDER})
    for candidate in candidates:
        existing = (
            get_by_mark(conn, candidate.product_kind, candidate.mark)
            if table_ready
            else None
        )
        if existing is not None and existing.match_status == MATCH_STATUS_MANUAL:
            report.skipped_manual += 1
            continue
        report.by_kind[candidate.product_kind] = (
            report.by_kind.get(candidate.product_kind, 0) + 1
        )
        if candidate.guid_1c_u:
            report.with_guid_u += 1
        if existing is None:
            report.inserted += 1
            if not dry_run:
                upsert(
                    conn,
                    candidate.product_kind,
                    candidate.mark,
                    guid_1c=candidate.guid_1c,
                    guid_1c_u=candidate.guid_1c_u,
                    match_status=MATCH_STATUS_AUTO,
                )
            continue
        if _row_unchanged(
            existing.guid_1c,
            existing.guid_1c_u,
            existing.match_status,
            candidate,
        ):
            continue
        report.updated += 1
        if not dry_run:
            upsert(
                conn,
                candidate.product_kind,
                candidate.mark,
                guid_1c=candidate.guid_1c,
                guid_1c_u=candidate.guid_1c_u,
                match_status=MATCH_STATUS_AUTO,
            )
    return report


def run_bootstrap(
    conn: sqlite3.Connection,
    name_idx,
    db: dict,
    *,
    dry_run: bool = False,
    groups: Iterable[str] = NON_PLATE_QUEUE_GROUPS,
) -> BootstrapReport:
    candidates, skipped_not_ready, skipped_dup, dup_labels = collect_auto_candidates(
        name_idx, db, groups=groups
    )
    report = apply_bootstrap(conn, candidates, dry_run=dry_run)
    report.skipped_not_ready = skipped_not_ready
    report.skipped_dup = skipped_dup
    report.skipped_dup_marks = dup_labels
    return report


def format_report(report: BootstrapReport) -> str:
    lines = [
        f"inserted: {report.inserted}",
        f"updated: {report.updated}",
        f"skipped_manual: {report.skipped_manual}",
        f"skipped_dup: {report.skipped_dup}",
        "by product_kind:",
    ]
    for kind in KIND_ORDER:
        lines.append(f"  {kind}: {report.by_kind.get(kind, 0)}")
    lines.append(f"skipped_not_ready: {report.skipped_not_ready}")
    lines.append(f"with_guid_u: {report.with_guid_u}")
    if report.skipped_dup_marks:
        lines.append("skipped_dup_marks:")
        for label in report.skipped_dup_marks:
            lines.append(f"  {label}")
    return "\n".join(lines)


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Залить однозначные GUID не-плит в nomenclature_guid (status=auto)."
    )
    parser.add_argument(
        "--db",
        default=str(PRICE_DB_PATH),
        help=f"Путь к pb.db (по умолчанию: {PRICE_DB_PATH})",
    )
    parser.add_argument(
        "--guid-dir",
        default=str(DEFAULT_GUID_DIR),
        help="Папка выгрузок 1С (по умолчанию: банк знаний/Новая папка/GUID)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Посчитать inserted/updated без записи в БД",
    )
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    guid_dir = Path(args.guid_dir)
    db_path = Path(args.db)
    if not guid_dir.is_dir():
        print(f"❌ Нет папки выгрузок 1С: {guid_dir}")
        return 1
    if not db_path.is_file():
        print(f"❌ Нет pb.db: {db_path}")
        return 1
    for fname in PRICE_FILE_NAMES.values():
        path = guid_dir / fname
        if not path.is_file():
            print(f"❌ Нет файла выгрузки: {path}")
            return 1

    print("Загрузка выгрузок 1С и прайса pb.db…")
    _c1c, name_idx = load_1c(guid_dir)
    catalog = load_db(db_path)
    conn = sqlite3.connect(str(db_path))
    try:
        report = run_bootstrap(conn, name_idx, catalog, dry_run=args.dry_run)
        if args.dry_run:
            print("dry-run (запись не выполнялась)")
        else:
            conn.commit()
            print("commit")
        print(format_report(report))
        auto_total = sum(report.by_kind.get(kind, 0) for kind in KIND_ORDER)
        print(f"auto candidates: {auto_total}")
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
