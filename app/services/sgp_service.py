"""SGP warehouse service: list / unlink / relink / free plates / progress."""

from __future__ import annotations

import sqlite3
from typing import Any, Literal

from app.domain.enums import (
    PlateStatus,
    PlateTransitionReason,
)
from app.schemas.sgp import (
    SgpFreePlateItem,
    SgpFreePlatesResponse,
    SgpMutationResponse,
    SgpPlateItem,
    SgpPlatesResponse,
    SgpProgress,
)
from core.kp_db_audit import audit_append
from core.kp_db_common import _connect
from core.kp_db_plates_completion import (
    check_and_update_kp_completion,
)
from core.kp_db_schema import ensure_schema
from core.kp_db_shipments import (
    allocated_qty_for_completed_plate,
    open_reservation_for_completed_plate,
)

FilterKind = Literal["all", "linked", "unlinked"]


class SgpError(ValueError):
    """Domain validation error for SGP operations (maps to 422)."""

    def __init__(self, message: str, *, code: str = "sgp_error") -> None:
        super().__init__(message)
        self.code = code


def _assert_not_allocated(cur: sqlite3.Cursor, sgp_id: int, row_qty: int, qty: int) -> None:
    """SHIP-205: unlink/relink оперируют только свободной частью строки СГП.

    Часть, зарезервированная open-рейсом, недоступна → 422 ``sgp_row_allocated``.
    """
    allocated = allocated_qty_for_completed_plate(cur, sgp_id)
    if allocated <= 0 or qty <= row_qty - allocated:
        return
    reservation = open_reservation_for_completed_plate(cur, sgp_id)
    if reservation:
        shipment_id, shipment_date, _reserved = reservation
        where = f"рейсом #{shipment_id} от {shipment_date}"
    else:
        where = "открытым рейсом"
    raise SgpError(
        f"Плита зарезервирована {where}: свободно {row_qty - allocated} из {row_qty}, "
        f"запрошено {qty}",
        code="sgp_row_allocated",
    )


class SgpService:
    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def list_plates(self, *, filter: FilterKind = "all") -> SgpPlatesResponse:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            where = ""
            if filter == "linked":
                where = "WHERE cp.kp_id IS NOT NULL"
            elif filter == "unlinked":
                where = "WHERE cp.kp_id IS NULL"
            cur.execute(
                f"""
                SELECT
                    cp.id, cp.kp_id, cp.plate_name, cp.length_m, cp.width_m,
                    cp.load_class, cp.qty, cp.completed_date, cp.production_day,
                    cp.plan_id, cp.nomenclature_id,
                    o.customer_name, o.execution_terms
                FROM completed_plates cp
                LEFT JOIN KP_offers o ON o.kp_id = cp.kp_id
                {where}
                ORDER BY cp.id
                """
            )
            rows = cur.fetchall()
            items: list[SgpPlateItem] = []
            progress_cache: dict[int, SgpProgress] = {}
            for row in rows:
                progress = None
                kp_id = row["kp_id"]
                if kp_id is not None:
                    kid = int(kp_id)
                    if kid not in progress_cache:
                        progress_cache[kid] = self._sgp_progress_on_cursor(cur, kid)
                    progress = progress_cache[kid]
                items.append(
                    SgpPlateItem(
                        id=int(row["id"]),
                        kp_id=int(kp_id) if kp_id is not None else None,
                        plate_name=str(row["plate_name"] or ""),
                        length_m=row["length_m"],
                        width_m=row["width_m"],
                        load_class=row["load_class"],
                        qty=int(row["qty"] or 0),
                        completed_date=row["completed_date"],
                        production_day=row["production_day"],
                        plan_id=row["plan_id"],
                        nomenclature_id=row["nomenclature_id"],
                        customer_name=row["customer_name"],
                        execution_terms=row["execution_terms"],
                        sgp_progress=progress,
                    )
                )
            return SgpPlatesResponse(items=items, count=len(items), filter=filter)
        finally:
            conn.close()

    def free_plates(
        self,
        *,
        plate_name: str | None = None,
        length_m: float | None = None,
        width_m: float | None = None,
        load_class: int | None = None,
    ) -> SgpFreePlatesResponse:
        """Unlinked SGP plates, optionally filtered by strict identity (wizard hints)."""
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            clauses = ["kp_id IS NULL", "qty > 0"]
            params: list[Any] = []
            if plate_name:
                clauses.append("plate_name = ?")
                params.append(plate_name)
            if length_m is not None:
                clauses.append("ABS(length_m - ?) < 0.005")
                params.append(float(length_m))
            if width_m is not None:
                clauses.append("ABS(width_m - ?) < 0.005")
                params.append(float(width_m))
            if load_class is not None:
                clauses.append("load_class = ?")
                params.append(int(load_class))
            where = " AND ".join(clauses)
            cur.execute(
                f"""
                SELECT id, plate_name, length_m, width_m, load_class, qty, completed_date
                FROM completed_plates
                WHERE {where}
                ORDER BY completed_date ASC, id ASC
                """,
                params,
            )
            items = [
                SgpFreePlateItem(
                    id=int(r["id"]),
                    plate_name=str(r["plate_name"] or ""),
                    length_m=r["length_m"],
                    width_m=r["width_m"],
                    load_class=r["load_class"],
                    qty=int(r["qty"] or 0),
                    completed_date=r["completed_date"],
                )
                for r in cur.fetchall()
            ]
            return SgpFreePlatesResponse(items=items, count=len(items))
        finally:
            conn.close()

    def reduce_selected_qty_for_reservations(
        self,
        *,
        selected_plate_qty: dict[int, dict[int, int]] | None,
        sgp_reservations: list[dict[str, Any]],
    ) -> dict[int, dict[int, int]]:
        """Reduce optimizer demand by reserved SGP qty (strict match, FIFO by plate id).

        Returns a new ``selected_plate_qty`` map. Plates without an override keep
        full DB qty in the loader; only affected rows get explicit reduced qty.
        """
        if not sgp_reservations:
            return dict(selected_plate_qty or {})

        ensure_schema(self.db_path)
        result: dict[int, dict[int, int]] = {
            int(kp): {int(pid): int(q) for pid, q in plates.items()}
            for kp, plates in (selected_plate_qty or {}).items()
        }

        conn = _connect(self.db_path)
        try:
            cur = conn.cursor()
            for item in sgp_reservations:
                sgp_id = int(item["sgp_id"])
                target_kp_id = int(item["target_kp_id"])
                need = int(item["qty"])
                if need <= 0:
                    continue

                cur.execute(
                    """
                    SELECT plate_name, length_m, width_m, load_class, qty
                    FROM completed_plates
                    WHERE id = ? AND kp_id IS NULL AND qty > 0
                    """,
                    (sgp_id,),
                )
                sgp_row = cur.fetchone()
                if not sgp_row:
                    raise SgpError(
                        f"Свободная плита СГП #{sgp_id} не найдена",
                        code="sgp_not_found",
                    )
                plate_name, length_m, width_m, load_class, free_qty = sgp_row
                need = min(need, int(free_qty))

                cur.execute(
                    """
                    SELECT id, qty FROM kp_plates
                    WHERE kp_id = ?
                      AND plate_name = ?
                      AND ABS(COALESCE(length_m, 0) - ?) < 0.005
                      AND ABS(COALESCE(width_m, 0) - ?) < 0.005
                      AND COALESCE(load_class, 0) = ?
                      AND status = 'в производстве'
                      AND qty > 0
                    ORDER BY id
                    """,
                    (
                        target_kp_id,
                        plate_name,
                        float(length_m or 0),
                        float(width_m or 0),
                        int(load_class or 0),
                    ),
                )
                demand_rows = cur.fetchall()
                if not demand_rows:
                    raise SgpError(
                        "Нет открытой потребности для резерва со СГП",
                        code="sgp_no_matching_demand",
                    )

                remaining = need
                kp_map = result.setdefault(target_kp_id, {})
                for plate_id, db_qty in demand_rows:
                    if remaining <= 0:
                        break
                    pid = int(plate_id)
                    base = kp_map.get(pid, int(db_qty))
                    take = min(remaining, base)
                    if take <= 0:
                        continue
                    new_qty = base - take
                    if new_qty > 0:
                        kp_map[pid] = new_qty
                    else:
                        # Keep explicit 0 so loader skips (qty <= 0 continue)
                        kp_map[pid] = 0
                    remaining -= take

                if remaining > 0:
                    raise SgpError(
                        f"Недостаточно потребности для резерва {need} шт "
                        f"(не хватает {remaining})",
                        code="sgp_no_matching_demand",
                    )
            return result
        finally:
            conn.close()

    def build_from_sgp_rows(
        self,
        sgp_reservations: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        """Enrich reservations with plate identity for plan/day UI."""
        if not sgp_reservations:
            return []
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            rows: list[dict[str, Any]] = []
            for item in sgp_reservations:
                sgp_id = int(item["sgp_id"])
                cur.execute(
                    """
                    SELECT id, plate_name, length_m, width_m, load_class, qty
                    FROM completed_plates WHERE id = ?
                    """,
                    (sgp_id,),
                )
                r = cur.fetchone()
                rows.append(
                    {
                        "sgp_id": sgp_id,
                        "target_kp_id": int(item["target_kp_id"]),
                        "qty": int(item["qty"]),
                        "plate_name": (r["plate_name"] if r else "") or "",
                        "length_m": r["length_m"] if r else None,
                        "width_m": r["width_m"] if r else None,
                        "load_class": r["load_class"] if r else None,
                        "source": "с СГП",
                    }
                )
            return rows
        finally:
            conn.close()

    def export_plan_sgp_xlsx(self, plan_id: str, plan_payload: dict[str, Any]) -> bytes:
        """Build XLSX bytes for plan positions closed from SGP."""
        from io import BytesIO

        from openpyxl import Workbook
        from openpyxl.styles import Font

        reservations = list(plan_payload.get("sgp_reservations") or [])
        # Prefer enriched from_sgp if present; else rebuild from reservations.
        rows = list(plan_payload.get("from_sgp") or [])
        if not rows and reservations:
            rows = self.build_from_sgp_rows(reservations)

        wb = Workbook()
        ws = wb.active
        ws.title = "Со склада"
        headers = [
            "КП",
            "Плита",
            "Длина, м",
            "Ширина, м",
            "Нагрузка",
            "Qty",
            "SGP id",
            "Источник",
        ]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for row in rows:
            ws.append(
                [
                    row.get("target_kp_id") or row.get("kp_id"),
                    row.get("plate_name") or "",
                    row.get("length_m"),
                    row.get("width_m"),
                    row.get("load_class"),
                    row.get("qty"),
                    row.get("sgp_id"),
                    row.get("source") or "с СГП",
                ]
            )
        if not rows:
            ws.append(["—", "Нет позиций, закрытых со СГП", "", "", "", "", "", ""])

        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def sgp_progress(self, kp_id: int) -> SgpProgress:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            cur = conn.cursor()
            return self._sgp_progress_on_cursor(cur, kp_id)
        finally:
            conn.close()

    def unlink(
        self,
        sgp_id: int,
        qty: int,
        *,
        actor: str | None = None,
    ) -> SgpMutationResponse:
        if qty <= 0:
            raise SgpError("qty must be >= 1", code="sgp_invalid_qty")
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.cursor()
            cur.execute(
                """
                SELECT id, kp_id, plate_name, length_m, width_m, load_class, qty,
                       completed_date, production_day, nomenclature_id, plan_id
                FROM completed_plates WHERE id = ?
                """,
                (sgp_id,),
            )
            row = cur.fetchone()
            if not row:
                raise SgpError("Плита на СГП не найдена", code="sgp_not_found")
            (
                _id,
                kp_id,
                plate_name,
                length_m,
                width_m,
                load_class,
                row_qty,
                completed_date,
                production_day,
                nomenclature_id,
                plan_id,
            ) = row
            if kp_id is None:
                raise SgpError(
                    "Плита уже отвязана от КП",
                    code="sgp_already_unlinked",
                )
            row_qty = int(row_qty)
            if qty > row_qty:
                raise SgpError(
                    f"Нельзя отвязать {qty} из {row_qty}",
                    code="sgp_qty_exceeds",
                )
            _assert_not_allocated(cur, sgp_id, row_qty, qty)

            source_kp_id = int(kp_id)

            if qty < row_qty:
                # Split: keep linked remainder on original row, insert unlinked part
                cur.execute(
                    "UPDATE completed_plates SET qty = qty - ? WHERE id = ?",
                    (qty, sgp_id),
                )
                cur.execute(
                    """
                    INSERT INTO completed_plates (
                        kp_id, plate_name, length_m, width_m, load_class,
                        qty, completed_date, production_day, nomenclature_id, plan_id
                    ) VALUES (NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        plate_name,
                        length_m,
                        width_m,
                        load_class,
                        qty,
                        completed_date,
                        production_day,
                        nomenclature_id,
                        plan_id,
                    ),
                )
                free_id = int(cur.lastrowid)
            else:
                cur.execute(
                    "UPDATE completed_plates SET kp_id = NULL WHERE id = ?",
                    (sgp_id,),
                )
                free_id = sgp_id

            # Return demand to source KP as «в производстве»
            cur.execute(
                """
                SELECT id FROM kp_plates
                WHERE kp_id = ?
                  AND plate_name = ?
                  AND ABS(COALESCE(length_m, 0) - ?) < 0.005
                  AND ABS(COALESCE(width_m, 0) - ?) < 0.005
                  AND COALESCE(load_class, 0) = ?
                  AND status = 'в производстве'
                  AND (plan_id IS NULL OR plan_id = '')
                ORDER BY id LIMIT 1
                """,
                (
                    source_kp_id,
                    plate_name,
                    float(length_m or 0),
                    float(width_m or 0),
                    int(load_class or 0),
                ),
            )
            existing = cur.fetchone()
            if existing:
                cur.execute(
                    "UPDATE kp_plates SET qty = qty + ? WHERE id = ?",
                    (qty, int(existing[0])),
                )
                demand_plate_id = int(existing[0])
            else:
                # Clone identity from a historical kp_plates row via completed dims;
                # insert a fresh in-production row.
                cur.execute(
                    """
                    INSERT INTO kp_plates (
                        kp_id, position_number, plate_name, length_m, width_m,
                        load_class, qty, status, nomenclature_id
                    ) VALUES (?, 1, ?, ?, ?, ?, ?, 'в производстве', ?)
                    """,
                    (
                        source_kp_id,
                        plate_name,
                        length_m,
                        width_m,
                        load_class,
                        qty,
                        nomenclature_id,
                    ),
                )
                demand_plate_id = int(cur.lastrowid)

            audit_append(
                cur,
                plate_id=demand_plate_id,
                kp_id=source_kp_id,
                plate_name=plate_name,
                plan_id=plan_id,
                day_number=production_day,
                from_status=PlateStatus.ON_SGP.value,
                to_status=PlateStatus.IN_PRODUCTION.value,
                qty=qty,
                reason=PlateTransitionReason.SGP_UNLINK.value,
                actor=actor,
            )

            check_and_update_kp_completion(
                source_kp_id, self.db_path, _external_conn=conn
            )
            conn.commit()
            return SgpMutationResponse(
                ok=True,
                sgp_id=free_id,
                qty=qty,
                kp_id=None,
                message=f"Отвязано {qty} шт от КП #{source_kp_id}",
            )
        except SgpError:
            conn.rollback()
            raise
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def relink(
        self,
        sgp_id: int,
        *,
        target_kp_id: int,
        qty: int,
        actor: str | None = None,
    ) -> SgpMutationResponse:
        if qty <= 0:
            raise SgpError("qty must be >= 1", code="sgp_invalid_qty")
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.cursor()
            cur.execute(
                """
                SELECT id, kp_id, plate_name, length_m, width_m, load_class, qty,
                       completed_date, production_day, nomenclature_id, plan_id
                FROM completed_plates WHERE id = ?
                """,
                (sgp_id,),
            )
            row = cur.fetchone()
            if not row:
                raise SgpError("Плита на СГП не найдена", code="sgp_not_found")
            (
                _id,
                kp_id,
                plate_name,
                length_m,
                width_m,
                load_class,
                row_qty,
                completed_date,
                production_day,
                nomenclature_id,
                plan_id,
            ) = row
            if kp_id is not None:
                raise SgpError(
                    "Сначала отвяжите плиту от текущего КП",
                    code="sgp_must_unlink_first",
                )
            row_qty = int(row_qty)
            if qty > row_qty:
                raise SgpError(
                    f"Нельзя перепривязать {qty} из {row_qty}",
                    code="sgp_qty_exceeds",
                )
            _assert_not_allocated(cur, sgp_id, row_qty, qty)

            # Strict match open demand on target KP
            cur.execute(
                """
                SELECT id, qty FROM kp_plates
                WHERE kp_id = ?
                  AND plate_name = ?
                  AND ABS(COALESCE(length_m, 0) - ?) < 0.005
                  AND ABS(COALESCE(width_m, 0) - ?) < 0.005
                  AND COALESCE(load_class, 0) = ?
                  AND status = 'в производстве'
                  AND qty > 0
                ORDER BY id
                """,
                (
                    int(target_kp_id),
                    plate_name,
                    float(length_m or 0),
                    float(width_m or 0),
                    int(load_class or 0),
                ),
            )
            demand_rows = cur.fetchall()
            open_demand = sum(int(r[1]) for r in demand_rows)
            if open_demand <= 0:
                raise SgpError(
                    "У целевого КП нет открытой потребности с теми же параметрами плиты",
                    code="sgp_no_matching_demand",
                )
            if qty > open_demand:
                raise SgpError(
                    f"Открытая потребность {open_demand} шт, запрошено {qty}",
                    code="sgp_no_matching_demand",
                )

            # Deduct demand from target kp_plates
            remaining = qty
            for dem_id, dem_qty in demand_rows:
                if remaining <= 0:
                    break
                take = min(remaining, int(dem_qty))
                if take >= int(dem_qty):
                    cur.execute("DELETE FROM kp_plates WHERE id = ?", (int(dem_id),))
                else:
                    cur.execute(
                        "UPDATE kp_plates SET qty = qty - ? WHERE id = ?",
                        (take, int(dem_id)),
                    )
                remaining -= take

            # Bind SGP row (split if partial)
            if qty < row_qty:
                cur.execute(
                    "UPDATE completed_plates SET qty = qty - ? WHERE id = ?",
                    (qty, sgp_id),
                )
                cur.execute(
                    """
                    INSERT INTO completed_plates (
                        kp_id, plate_name, length_m, width_m, load_class,
                        qty, completed_date, production_day, nomenclature_id, plan_id
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        int(target_kp_id),
                        plate_name,
                        length_m,
                        width_m,
                        load_class,
                        qty,
                        completed_date,
                        production_day,
                        nomenclature_id,
                        plan_id,
                    ),
                )
                bound_id = int(cur.lastrowid)
            else:
                cur.execute(
                    "UPDATE completed_plates SET kp_id = ? WHERE id = ?",
                    (int(target_kp_id), sgp_id),
                )
                bound_id = sgp_id

            audit_append(
                cur,
                plate_id=None,
                kp_id=int(target_kp_id),
                plate_name=plate_name,
                plan_id=plan_id,
                day_number=production_day,
                from_status=PlateStatus.ON_SGP.value,
                to_status=PlateStatus.ON_SGP.value,
                qty=qty,
                reason=PlateTransitionReason.SGP_RELINK.value,
                actor=actor,
            )

            check_and_update_kp_completion(
                int(target_kp_id), self.db_path, _external_conn=conn
            )
            conn.commit()
            return SgpMutationResponse(
                ok=True,
                sgp_id=bound_id,
                qty=qty,
                kp_id=int(target_kp_id),
                target_kp_id=int(target_kp_id),
                message=f"Перепривязано {qty} шт к КП #{target_kp_id}",
            )
        except SgpError:
            conn.rollback()
            raise
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def reserve_on_conn(
        self,
        cur: sqlite3.Cursor,
        conn: sqlite3.Connection,
        *,
        sgp_id: int,
        target_kp_id: int,
        qty: int,
        plan_id: str | None = None,
        actor: str | None = None,
    ) -> int:
        """Atomic reserve inside an existing transaction (buildPlan)."""
        cur.execute(
            """
            SELECT id, kp_id, plate_name, length_m, width_m, load_class, qty,
                   completed_date, production_day, nomenclature_id, plan_id
            FROM completed_plates WHERE id = ?
            """,
            (sgp_id,),
        )
        row = cur.fetchone()
        if not row:
            raise SgpError("Плита на СГП не найдена", code="sgp_not_found")
        (
            _id,
            kp_id,
            plate_name,
            length_m,
            width_m,
            load_class,
            row_qty,
            completed_date,
            production_day,
            nomenclature_id,
            existing_plan_id,
        ) = row
        if kp_id is not None:
            raise SgpError("Плита уже привязана к КП", code="sgp_not_free")
        row_qty = int(row_qty)
        take = min(qty, row_qty)
        if take <= 0:
            return 0

        # Deduct open demand
        cur.execute(
            """
            SELECT id, qty FROM kp_plates
            WHERE kp_id = ?
              AND plate_name = ?
              AND ABS(COALESCE(length_m, 0) - ?) < 0.005
              AND ABS(COALESCE(width_m, 0) - ?) < 0.005
              AND COALESCE(load_class, 0) = ?
              AND status = 'в производстве'
              AND qty > 0
            ORDER BY id
            """,
            (
                int(target_kp_id),
                plate_name,
                float(length_m or 0),
                float(width_m or 0),
                int(load_class or 0),
            ),
        )
        demand_rows = cur.fetchall()
        open_demand = sum(int(r[1]) for r in demand_rows)
        take = min(take, open_demand)
        if take <= 0:
            raise SgpError(
                "Нет открытой потребности для резерва со СГП",
                code="sgp_no_matching_demand",
            )

        remaining = take
        for dem_id, dem_qty in demand_rows:
            if remaining <= 0:
                break
            chunk = min(remaining, int(dem_qty))
            if chunk >= int(dem_qty):
                cur.execute("DELETE FROM kp_plates WHERE id = ?", (int(dem_id),))
            else:
                cur.execute(
                    "UPDATE kp_plates SET qty = qty - ? WHERE id = ?",
                    (chunk, int(dem_id)),
                )
            remaining -= chunk

        bind_plan = plan_id or existing_plan_id
        if take < row_qty:
            cur.execute(
                "UPDATE completed_plates SET qty = qty - ? WHERE id = ?",
                (take, sgp_id),
            )
            cur.execute(
                """
                INSERT INTO completed_plates (
                    kp_id, plate_name, length_m, width_m, load_class,
                    qty, completed_date, production_day, nomenclature_id, plan_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(target_kp_id),
                    plate_name,
                    length_m,
                    width_m,
                    load_class,
                    take,
                    completed_date,
                    production_day,
                    nomenclature_id,
                    bind_plan,
                ),
            )
        else:
            cur.execute(
                "UPDATE completed_plates SET kp_id = ?, plan_id = COALESCE(?, plan_id) WHERE id = ?",
                (int(target_kp_id), bind_plan, sgp_id),
            )

        audit_append(
            cur,
            plate_id=None,
            kp_id=int(target_kp_id),
            plate_name=plate_name,
            plan_id=bind_plan,
            day_number=production_day,
            from_status=PlateStatus.ON_SGP.value,
            to_status=PlateStatus.ON_SGP.value,
            qty=take,
            reason=PlateTransitionReason.SGP_RESERVE.value,
            actor=actor,
        )
        check_and_update_kp_completion(
            int(target_kp_id), self.db_path, _external_conn=conn
        )
        return take

    def clear_plan_links(self, plan_id: str, *, _external_conn: sqlite3.Connection | None = None) -> int:
        """On delete_plan: nullify plan_id on SGP rows; do not touch qty."""
        own = _external_conn is None
        conn = _external_conn if not own else _connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute(
                "UPDATE completed_plates SET plan_id = NULL WHERE plan_id = ?",
                (plan_id,),
            )
            n = cur.rowcount
            if own:
                conn.commit()
            return int(n or 0)
        finally:
            if own:
                conn.close()

    @staticmethod
    def _sgp_progress_on_cursor(cur: sqlite3.Cursor, kp_id: int) -> SgpProgress:
        # Read-only: не freeze'им ordered_qty на GET/progress (freeze — только write-path).
        cur.execute(
            "SELECT ordered_qty FROM kp_meta WHERE kp_id = ?",
            (kp_id,),
        )
        meta = cur.fetchone()
        ordered = int(meta[0]) if meta and meta[0] is not None else None
        cur.execute(
            "SELECT COALESCE(SUM(qty), 0) FROM completed_plates WHERE kp_id = ?",
            (kp_id,),
        )
        n = int(cur.fetchone()[0] or 0)
        if ordered is None:
            cur.execute(
                "SELECT COALESCE(SUM(qty), 0) FROM kp_plates WHERE kp_id = ?",
                (kp_id,),
            )
            remaining = int(cur.fetchone()[0] or 0)
            ordered = remaining + n
        return SgpProgress(n=n, m=max(ordered, n))
