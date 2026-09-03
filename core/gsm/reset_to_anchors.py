"""Сброс ГСМ к imported-якорям: общая логика для CLI и HTTP API."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


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


@dataclass(frozen=True, slots=True)
class ResetResult:
    plan: ResetPlan
    backup_path: Path | None


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
    for anchor in plan.anchors:
        fuel = f"{anchor.fuel_end:.2f}" if anchor.fuel_end is not None else "—"
        odo = str(anchor.odometer_end) if anchor.odometer_end is not None else "—"
        lines.append(
            f"    v{anchor.vehicle_id} {anchor.name} {anchor.plate_number}: "
            f"wb#{anchor.waybill_id} {anchor.date} status={anchor.status} "
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
    keep_ids = [anchor.waybill_id for anchor in plan.anchors]
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


def run_reset(*, db_path: Path, apply: bool) -> ResetResult:
    if not db_path.is_file():
        raise ResetGsmError(f"БД не найдена: {db_path}")

    with _connect(db_path) as conn:
        plan = build_reset_plan(conn)
        if not apply:
            return ResetResult(plan=plan, backup_path=None)

    backup_path = backup_database(db_path)

    with _connect(db_path) as conn:
        plan = build_reset_plan(conn)
        apply_reset(conn, plan)
        after_wb = _count(conn, "gsm_waybill")
        after_tx = _count(conn, "gsm_transaction")
        after_batch = _count(conn, "gsm_import_batch")

    if after_wb != len(plan.anchors) or after_tx != 0 or after_batch != 0:
        raise ResetGsmError(
            f"после сброса неожиданные счётчики: "
            f"waybills={after_wb} txs={after_tx} batches={after_batch}"
        )
    return ResetResult(plan=plan, backup_path=backup_path)
