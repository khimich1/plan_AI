#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Разовый двусторонний ремонт сироты kp_plates.id=197 и призрака в JSON плана.

По умолчанию -- dry-run (ничего не пишет). Запись только с явным ``--apply``.
Живую БД не трогать без подтверждения пользователя и бэкапа.

Использование::

    .venv/bin/python scripts/repair_plan_orphan_197.py --db-path plita.db
    .venv/bin/python scripts/repair_plan_orphan_197.py --db-path plita.db --apply
"""
from __future__ import annotations

import argparse
import json
import logging
import sqlite3
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

DEFAULT_PLAN_ID = "plan_20260904_124248"
DEFAULT_PLATE_ID = 197
DEFAULT_TARGET_KP_ID = 4
DEFAULT_GHOST_KP_ID = 5
DEFAULT_DAY_NUMBER = 2

logger = logging.getLogger("repair_plan_orphan_197")


@dataclass(frozen=True)
class RepairConfig:
    plan_id: str = DEFAULT_PLAN_ID
    plate_id: int = DEFAULT_PLATE_ID
    target_kp_id: int = DEFAULT_TARGET_KP_ID
    ghost_kp_id: int = DEFAULT_GHOST_KP_ID
    day_number: int = DEFAULT_DAY_NUMBER


@dataclass
class RepairReport:
    dry_run: bool
    applied: bool
    already_fixed: bool = False
    changes: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Ремонт сироты id=197 (day_number=2) и призрака в JSON плана "
            f"{DEFAULT_PLAN_ID}. Без --apply ничего не пишет."
        )
    )
    parser.add_argument(
        "--db-path",
        default=None,
        help="Путь к SQLite. По умолчанию plita.db из настроек.",
    )
    parser.add_argument(
        "--plan-id",
        default=DEFAULT_PLAN_ID,
        help=f"ID плана (по умолчанию {DEFAULT_PLAN_ID}).",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Записать изменения. Без флага — только план (dry-run).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Явный dry-run (режим по умолчанию). Нельзя вместе с --apply.",
    )
    return parser.parse_args(argv)


def tracks_by_day_from_plan(plan: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    """Снимок tracks_by_day для гейта целостности (не мутирует план)."""
    out: dict[str, list[dict[str, Any]]] = {}
    for date_key, day_data in (plan.get("days") or {}).items():
        day_number = int((day_data or {}).get("day_number") or 0)
        tracks: list[dict[str, Any]] = []
        for track in (day_data or {}).get("tracks") or []:
            if not isinstance(track, dict):
                continue
            copied = dict(track)
            copied.setdefault("production_day", day_number)
            tracks.append(copied)
        out[str(date_key)] = tracks
    return out


def _iter_physical_items(plan: dict[str, Any]) -> Iterable[dict[str, Any]]:
    for day in (plan.get("days") or {}).values():
        if not isinstance(day, dict):
            continue
        for track in day.get("tracks") or []:
            if not isinstance(track, dict):
                continue
            for item in track.get("items") or []:
                if not isinstance(item, dict):
                    continue
                yield item
                for sec in item.get("secondary_cuts") or []:
                    if isinstance(sec, dict):
                        yield sec


def _pid_missing(item: dict[str, Any]) -> bool:
    pid = item.get("kp_plate_id")
    return pid is None or pid == "" or pid == 0


def _connect(db_path: str, *, apply: bool) -> sqlite3.Connection:
    if apply:
        conn = sqlite3.connect(db_path)
    else:
        uri = Path(db_path).resolve().as_posix()
        conn = sqlite3.connect(f"file:{uri}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def _find_ghost(
    plan: dict[str, Any],
    *,
    ghost_kp_id: int,
    plate_name: str,
    plate_id: int,
    target_kp_id: int,
) -> tuple[dict[str, Any] | None, dict[str, Any] | None, str | None]:
    """Возвращает (призрак, уже починенный item, ошибка)."""
    ghosts: list[dict[str, Any]] = []
    already: dict[str, Any] | None = None
    for item in _iter_physical_items(plan):
        name = str(item.get("plate_name") or item.get("label") or "")
        if plate_name and name and name != plate_name:
            continue
        if item.get("kp_plate_id") == plate_id and int(item.get("kp_id") or 0) == target_kp_id:
            already = item
            continue
        if int(item.get("kp_id") or 0) == ghost_kp_id and _pid_missing(item):
            ghosts.append(item)
    if len(ghosts) > 1:
        return None, already, (
            f"Найдено несколько призраков КП #{ghost_kp_id} без kp_plate_id "
            f"({len(ghosts)} шт.) — отказ, нужен ручной разбор."
        )
    ghost = ghosts[0] if ghosts else None
    return ghost, already, None


def repair_orphan_197(
    db_path: str,
    *,
    apply: bool = False,
    config: RepairConfig | None = None,
) -> RepairReport:
    cfg = config or RepairConfig()
    report = RepairReport(dry_run=not apply, applied=False)
    conn = _connect(db_path, apply=apply)
    try:
        row = conn.execute(
            "SELECT id, kp_id, plate_name, plan_id, status, day_number "
            "FROM kp_plates WHERE id = ?",
            (cfg.plate_id,),
        ).fetchone()
        plan_row = conn.execute(
            "SELECT payload_json, version FROM production_plans WHERE id = ?",
            (cfg.plan_id,),
        ).fetchone()
        if row is None:
            report.errors.append(f"Строка kp_plates.id={cfg.plate_id} не найдена.")
            return report
        if plan_row is None:
            report.errors.append(f"План {cfg.plan_id} не найден в production_plans.")
            return report
        if int(row["kp_id"]) != cfg.target_kp_id:
            report.errors.append(
                f"kp_plates.id={cfg.plate_id} принадлежит КП #{row['kp_id']}, "
                f"ожидался КП #{cfg.target_kp_id}."
            )
            return report
        if str(row["plan_id"] or "") != cfg.plan_id:
            report.errors.append(
                f"kp_plates.id={cfg.plate_id} plan_id={row['plan_id']!r}, "
                f"ожидался {cfg.plan_id}."
            )
            return report

        payload = json.loads(plan_row["payload_json"])
        if not isinstance(payload, dict):
            report.errors.append("payload_json плана не объект.")
            return report

        plate_name = str(row["plate_name"] or "")
        ghost, already_item, ghost_error = _find_ghost(
            payload,
            ghost_kp_id=cfg.ghost_kp_id,
            plate_name=plate_name,
            plate_id=cfg.plate_id,
            target_kp_id=cfg.target_kp_id,
        )
        if ghost_error:
            report.errors.append(ghost_error)
            return report

        row_day = row["day_number"]
        row_ok = row_day == cfg.day_number
        item_ok = already_item is not None and ghost is None
        if row_ok and item_ok:
            report.already_fixed = True
            report.changes.append(
                f"Уже исправлено: id={cfg.plate_id} day_number={cfg.day_number}, "
                f"item КП #{cfg.target_kp_id} kp_plate_id={cfg.plate_id}."
            )
            logger.info("%s", report.changes[-1])
            return report

        if not row_ok:
            report.changes.append(
                f"kp_plates.id={cfg.plate_id}: day_number {row_day!r} → {cfg.day_number}"
            )
        if ghost is not None:
            report.changes.append(
                f"item-призрак: kp_id {ghost.get('kp_id')}→{cfg.target_kp_id}, "
                f"kp_plate_id {ghost.get('kp_plate_id')!r}→{cfg.plate_id} "
                f"({plate_name})"
            )
        elif not item_ok:
            report.errors.append(
                f"Призрак не найден: нет item КП #{cfg.ghost_kp_id} без kp_plate_id "
                f"для {plate_name}."
            )
            return report

        for line in report.changes:
            logger.info("%s", line)

        if not apply:
            return report

        if ghost is not None:
            ghost["kp_id"] = cfg.target_kp_id
            ghost["kp_plate_id"] = cfg.plate_id
        conn.execute("BEGIN")
        conn.execute(
            "UPDATE kp_plates SET day_number = ? WHERE id = ?",
            (cfg.day_number, cfg.plate_id),
        )
        conn.execute(
            """
            UPDATE production_plans
            SET payload_json = ?, version = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (
                json.dumps(payload, ensure_ascii=False),
                int(plan_row["version"]) + 1,
                cfg.plan_id,
            ),
        )
        conn.commit()
        report.applied = True
        logger.info("Записано в %s", db_path)
        return report
    finally:
        conn.close()


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.apply and args.dry_run:
        print("Нельзя одновременно --apply и --dry-run.", file=sys.stderr)
        return 2
    from app.core.settings import get_settings

    db_path = args.db_path or str(get_settings().plita_db_path)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    config = RepairConfig(plan_id=args.plan_id)
    prefix = "DRY-RUN" if not args.apply else "APPLY"
    print(f"[{prefix}] БД: {db_path}; план: {config.plan_id}")
    report = repair_orphan_197(db_path, apply=bool(args.apply), config=config)
    for line in report.changes:
        print(f"  {line}")
    for err in report.errors:
        print(f"Ошибка: {err}", file=sys.stderr)
    if report.errors:
        return 1
    if report.dry_run:
        print(
            "Изменения не записаны. Для записи нужен явный --apply "
            "(только на бэкапе, после подтверждения)."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
