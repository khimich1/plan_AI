#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сброс ГСМ к импортированным якорям (подготовка к тестовому прогону).

Оставляет на каждую активную машину один последний ``gsm_waybill`` с
``source='imported'`` (статус ``exported``), удаляет остальные ПЛ,
все ``gsm_transaction`` и ``gsm_import_batch``.

Справочники (машины, водители, карты, маршруты, станции, настройки)
не трогает.

По умолчанию — dry-run (только отчёт). Запись: ``--apply``
(сначала sqlite backup API → ``*.bak-before-gsm-test-YYYYMMDD-HHMMSS``).

Пример:
  .venv/bin/python scripts/reset_gsm_to_anchors.py --db plita.db
  .venv/bin/python scripts/reset_gsm_to_anchors.py --db plita.db --apply
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.kp_db_common import DEFAULT_DB


class ResetGsmError(RuntimeError):
    """Нельзя безопасно сбросить БД к якорям."""


@dataclass(frozen=True, slots=True)
class AnchorRow:
    waybill_id: int
    vehicle_id: int
    name: str
    plate_number: str
    date: str
    status: str
    source: str
    odometer_end: int | None
    fuel_end: float | None


@dataclass(frozen=True, slots=True)
class ResetPlan:
    anchors: tuple[AnchorRow, ...]
    waybills_total: int
    waybills_to_delete: int
    txs_total: int
    batches_total: int
    routes_total: int
    cards_total: int


def _connect(db_path: str | Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _count(conn: sqlite3.Connection, table: str) -> int:
    row = conn.execute(f"SELECT COUNT(*) AS n FROM {table}").fetchone()
    return int(row["n"])


def collect_imported_anchors(conn: sqlite3.Connection) -> tuple[AnchorRow, ...]:
    """Последний imported ПЛ на каждую активную машину; иначе Raise."""
    vehicles = conn.execute(
        """
        SELECT id, name, plate_number
        FROM gsm_vehicle
        WHERE is_active = 1
        ORDER BY id
        """
    ).fetchall()
    if not vehicles:
        raise ResetGsmError("нет активных машин в gsm_vehicle")

    anchors: list[AnchorRow] = []
    missing: list[str] = []
    for vehicle in vehicles:
        row = conn.execute(
            """
            SELECT id, vehicle_id, date, status, source, odometer_end, fuel_end
            FROM gsm_waybill
            WHERE vehicle_id = ? AND source = 'imported'
            ORDER BY date DESC, id DESC
            LIMIT 1
            """,
            (int(vehicle["id"]),),
        ).fetchone()
        if row is None:
            missing.append(
                f"#{vehicle['id']} {vehicle['name']} ({vehicle['plate_number']})"
            )
            continue
        anchors.append(
            AnchorRow(
                waybill_id=int(row["id"]),
                vehicle_id=int(row["vehicle_id"]),
                name=str(vehicle["name"]),
                plate_number=str(vehicle["plate_number"]),
                date=str(row["date"]),
                status=str(row["status"]),
                source=str(row["source"]),
                odometer_end=(
                    int(row["odometer_end"]) if row["odometer_end"] is not None else None
                ),
                fuel_end=(
                    float(row["fuel_end"]) if row["fuel_end"] is not None else None
                ),
            )
        )

    if missing:
        raise ResetGsmError(
            "нет imported-якоря для машин: " + "; ".join(missing)
        )
    return tuple(anchors)


def build_reset_plan(conn: sqlite3.Connection) -> ResetPlan:
    anchors = collect_imported_anchors(conn)
    waybills_total = _count(conn, "gsm_waybill")
    return ResetPlan(
        anchors=anchors,
        waybills_total=waybills_total,
        waybills_to_delete=waybills_total - len(anchors),
        txs_total=_count(conn, "gsm_transaction"),
        batches_total=_count(conn, "gsm_import_batch"),
        routes_total=_count(conn, "gsm_route"),
        cards_total=_count(conn, "gsm_fuel_card"),
    )


def format_plan(plan: ResetPlan, *, apply: bool) -> str:
    mode = "APPLY" if apply else "DRY-RUN"
    lines = [
        f"[{mode}] сброс ГСМ к imported-якорям",
        f"  якорей оставить: {len(plan.anchors)}",
        f"  ПЛ удалить: {plan.waybills_to_delete} (из {plan.waybills_total})",
        f"  транзакций удалить: {plan.txs_total}",
        f"  батчей импорта удалить: {plan.batches_total}",
        f"  маршрутов (не трогаем): {plan.routes_total}",
        f"  карт (не трогаем): {plan.cards_total}",
        "  якоря:",
    ]
    for a in plan.anchors:
        fuel = f"{a.fuel_end:.2f}" if a.fuel_end is not None else "—"
        odo = str(a.odometer_end) if a.odometer_end is not None else "—"
        lines.append(
            f"    v{a.vehicle_id} {a.name} {a.plate_number}: "
            f"wb#{a.waybill_id} {a.date} status={a.status} "
            f"odo_end={odo} fuel_end={fuel}"
        )
    return "\n".join(lines)


def backup_database(db_path: Path) -> Path:
    """Полный снимок через sqlite backup API (безопасно при WAL)."""
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak_path = db_path.with_name(f"{db_path.name}.bak-before-gsm-test-{stamp}")
    if bak_path.exists():
        raise ResetGsmError(f"файл бэкапа уже существует: {bak_path}")

    src = sqlite3.connect(str(db_path))
    try:
        dst = sqlite3.connect(str(bak_path))
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    return bak_path


def apply_reset(conn: sqlite3.Connection, plan: ResetPlan) -> None:
    keep_ids = [a.waybill_id for a in plan.anchors]
    placeholders = ",".join("?" for _ in keep_ids)
    try:
        conn.execute("BEGIN")
        conn.execute(
            f"""
            UPDATE gsm_waybill
            SET status = 'exported'
            WHERE id IN ({placeholders})
            """,
            keep_ids,
        )
        conn.execute(
            f"""
            DELETE FROM gsm_waybill
            WHERE id NOT IN ({placeholders})
            """,
            keep_ids,
        )
        conn.execute("DELETE FROM gsm_transaction")
        conn.execute("DELETE FROM gsm_import_batch")
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def run_reset(*, db_path: Path, apply: bool) -> ResetPlan:
    if not db_path.is_file():
        raise ResetGsmError(f"БД не найдена: {db_path}")

    with _connect(db_path) as conn:
        plan = build_reset_plan(conn)
        print(format_plan(plan, apply=apply))
        if not apply:
            print("  (ничего не записано; передайте --apply для выполнения)")
            return plan

    bak = backup_database(db_path)
    print(f"  бэкап: {bak}")

    with _connect(db_path) as conn:
        # Пересобрать план после бэкапа (на случай гонки) и снова проверить якоря.
        plan = build_reset_plan(conn)
        apply_reset(conn, plan)
        after_wb = _count(conn, "gsm_waybill")
        after_tx = _count(conn, "gsm_transaction")
        after_batch = _count(conn, "gsm_import_batch")
        after_routes = _count(conn, "gsm_route")
        after_cards = _count(conn, "gsm_fuel_card")

    print(
        "  после сброса: "
        f"waybills={after_wb} txs={after_tx} batches={after_batch} "
        f"routes={after_routes} cards={after_cards}"
    )
    if after_wb != len(plan.anchors) or after_tx != 0 or after_batch != 0:
        raise ResetGsmError(
            f"после сброса неожиданные счётчики: "
            f"waybills={after_wb} txs={after_tx} batches={after_batch}"
        )
    return plan


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Сброс ГСМ к последним imported-якорям: удаляет прочие ПЛ, "
            "транзакции и батчи импорта. По умолчанию dry-run."
        )
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=Path(DEFAULT_DB),
        help=f"Путь к SQLite (по умолчанию: {DEFAULT_DB})",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Выполнить сброс (с бэкапом). Без флага — только отчёт.",
    )
    return parser.parse_args(list(argv) if argv is not None else None)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        run_reset(db_path=args.db.resolve(), apply=bool(args.apply))
    except ResetGsmError as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
