"""Read-only KP readiness aggregation for manager view (A — готовность КП)."""

from __future__ import annotations

import sqlite3
from datetime import datetime

from app.domain.enums import KpStatus, PlateStatus
from app.planning.plan_manager import get_plan_day_to_date_mapping
from app.schemas.archive import (
    KpReadinessPositionItem,
    KpReadinessStep,
    KpReadinessStepState,
    KpReadinessSummary,
)
from app.schemas.sgp import SgpProgress
from app.services.sgp_service import SgpService
from core.kp.offers_read import get_kp_completion_percentage
from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema

_READINESS_STATUSES = frozenset({KpStatus.IN_WORK.value, KpStatus.ON_SGP.value})

_STEP_LABELS: dict[str, str] = {
    "kp": "КП",
    "production": "Производство",
    "sgp": "СГП",
    "release": "Выдача",
    "closed": "Закрыто",
}

_RELEASE_NOTE = "Выдача с СГП — в следующем обновлении"


class KpReadinessService:
    """Read-only aggregation of kp_plates + completed_plates for manager readiness view."""

    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def list_positions(self, kp_id: int, *, status: str | None = None) -> list[KpReadinessPositionItem]:
        if status is not None and status not in _READINESS_STATUSES:
            return []
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute(
                """
                WITH identities AS (
                    SELECT plate_name, length_m, width_m, load_class
                    FROM kp_plates
                    WHERE kp_id = ?
                    UNION
                    SELECT plate_name, length_m, width_m, load_class
                    FROM completed_plates
                    WHERE kp_id = ?
                ),
                grouped AS (
                    SELECT
                        plate_name,
                        ROUND(COALESCE(length_m, 0), 3) AS length_key,
                        ROUND(COALESCE(width_m, 0), 3) AS width_key,
                        load_class,
                        MIN(length_m) AS length_m,
                        MIN(width_m) AS width_m
                    FROM identities
                    GROUP BY plate_name, length_key, width_key, load_class
                )
                SELECT
                    g.plate_name,
                    g.length_m,
                    g.width_m,
                    g.load_class,
                    (
                        SELECT MIN(kp.position_number)
                        FROM kp_plates kp
                        WHERE kp.kp_id = ?
                          AND kp.plate_name = g.plate_name
                          AND kp.load_class IS g.load_class
                          AND ABS(COALESCE(kp.length_m, 0) - g.length_key) < 0.005
                          AND ABS(COALESCE(kp.width_m, 0) - g.width_key) < 0.005
                    ) AS position_number,
                    COALESCE((
                        SELECT SUM(kp.qty)
                        FROM kp_plates kp
                        WHERE kp.kp_id = ?
                          AND kp.status = ?
                          AND kp.plate_name = g.plate_name
                          AND kp.load_class IS g.load_class
                          AND ABS(COALESCE(kp.length_m, 0) - g.length_key) < 0.005
                          AND ABS(COALESCE(kp.width_m, 0) - g.width_key) < 0.005
                    ), 0) AS in_plan,
                    COALESCE((
                        SELECT SUM(kp.qty)
                        FROM kp_plates kp
                        WHERE kp.kp_id = ?
                          AND kp.status = ?
                          AND kp.plate_name = g.plate_name
                          AND kp.load_class IS g.load_class
                          AND ABS(COALESCE(kp.length_m, 0) - g.length_key) < 0.005
                          AND ABS(COALESCE(kp.width_m, 0) - g.width_key) < 0.005
                    ), 0) AS remaining,
                    COALESCE((
                        SELECT SUM(cp.qty)
                        FROM completed_plates cp
                        WHERE cp.kp_id = ?
                          AND cp.plate_name = g.plate_name
                          AND cp.load_class IS g.load_class
                          AND ABS(COALESCE(cp.length_m, 0) - g.length_key) < 0.005
                          AND ABS(COALESCE(cp.width_m, 0) - g.width_key) < 0.005
                    ), 0) AS on_sgp
                FROM grouped g
                """,
                (
                    kp_id,
                    kp_id,
                    kp_id,
                    kp_id,
                    PlateStatus.IN_PLAN.value,
                    kp_id,
                    PlateStatus.IN_PRODUCTION.value,
                    kp_id,
                ),
            )
            items: list[KpReadinessPositionItem] = []
            for row in cur.fetchall():
                in_plan = int(row["in_plan"] or 0)
                remaining = int(row["remaining"] or 0)
                on_sgp = int(row["on_sgp"] or 0)
                ordered = in_plan + remaining + on_sgp
                plate_name = str(row["plate_name"] or "")
                items.append(
                    KpReadinessPositionItem(
                        position_number=row["position_number"],
                        plate_name=plate_name,
                        length_m=row["length_m"],
                        width_m=row["width_m"],
                        load_class=row["load_class"],
                        label=plate_name,
                        ordered=ordered,
                        in_plan=in_plan,
                        on_sgp=on_sgp,
                        remaining=remaining,
                    )
                )
            items.sort(key=_position_sort_key)
            return items
        finally:
            conn.close()

    def build_summary(self, kp_id: int, *, status: str) -> KpReadinessSummary | None:
        if status not in _READINESS_STATUSES:
            return None

        progress = SgpService(db_path=self.db_path).sgp_progress(kp_id)
        completion = get_kp_completion_percentage(kp_id, self.db_path)
        in_production_qty = int(completion.get("in_production") or 0)
        completion_pct = float(completion.get("percentage") or 0.0)
        issuable_qty = progress.n

        positions = self.list_positions(kp_id)
        in_plan_total = sum(p.in_plan for p in positions)
        remaining_total = sum(p.remaining for p in positions)

        summary_text = _format_summary(progress, in_production_qty)
        client_copy_text = _format_client_copy(kp_id, progress, in_production_qty)

        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            plan_mapping_cache: dict[str, dict[int, str]] = {}
            expected_sgp_date, expected_sgp_date_label, fully_scheduled = self._resolve_expected_sgp_date(
                cur,
                kp_id,
                remaining_total=remaining_total,
                progress=progress,
                in_plan_total=in_plan_total,
                plan_mapping_cache=plan_mapping_cache,
            )
        finally:
            conn.close()

        if expected_sgp_date_label:
            client_copy_text = _append_expected_date_to_copy(client_copy_text, expected_sgp_date_label)

        return KpReadinessSummary(
            completion_percentage=completion_pct,
            sgp_progress=progress,
            issuable_qty=issuable_qty,
            in_production_qty=in_production_qty,
            summary_text=summary_text,
            client_copy_text=client_copy_text,
            steps=_build_steps(
                progress=progress,
                completion_pct=completion_pct,
                in_plan_total=in_plan_total,
                remaining_total=remaining_total,
            ),
            release_note=_RELEASE_NOTE,
            expected_sgp_date=expected_sgp_date,
            expected_sgp_date_label=expected_sgp_date_label,
            fully_scheduled=fully_scheduled,
        )

    def _resolve_expected_sgp_date(
        self,
        cur: sqlite3.Cursor,
        kp_id: int,
        *,
        remaining_total: int,
        progress: SgpProgress,
        in_plan_total: int,
        plan_mapping_cache: dict[str, dict[int, str]],
    ) -> tuple[str | None, str | None, bool]:
        if remaining_total > 0 or progress.n >= progress.m or in_plan_total <= 0:
            return None, None, False

        cur.execute(
            """
            SELECT DISTINCT plan_id, day_number
            FROM kp_plates
            WHERE kp_id = ?
              AND status = ?
              AND plan_id IS NOT NULL
              AND day_number IS NOT NULL
            """,
            (kp_id, PlateStatus.IN_PLAN.value),
        )

        resolved_dates: list[str] = []
        for row in cur.fetchall():
            plan_id = str(row["plan_id"])
            day_number = int(row["day_number"])
            if plan_id not in plan_mapping_cache:
                try:
                    plan_mapping_cache[plan_id] = get_plan_day_to_date_mapping(plan_id)
                except Exception:
                    plan_mapping_cache[plan_id] = {}
            calendar_date = plan_mapping_cache[plan_id].get(day_number)
            if calendar_date:
                resolved_dates.append(str(calendar_date))

        if not resolved_dates:
            return None, None, False

        iso_date = max(resolved_dates)
        label = datetime.strptime(iso_date, "%Y-%m-%d").strftime("%d.%m.%Y")
        return iso_date, label, True


def _position_sort_key(item: KpReadinessPositionItem) -> tuple[int, str]:
    pos = item.position_number
    if pos is None:
        return (1, item.label)
    return (0, f"{pos:08d}_{item.label}")


def _format_summary(progress: SgpProgress, in_production_qty: int) -> str:
    n, m = progress.n, progress.m
    if n > 0 and in_production_qty > 0:
        return (
            f"{n} из {m} шт на складе, {in_production_qty} в производстве. "
            f"Можно выдать {n} шт."
        )
    if n > 0 and in_production_qty == 0:
        return f"{n} из {m} шт на складе. Можно выдать {n} шт."
    if n == 0 and in_production_qty > 0:
        return f"Заказ в производстве ({in_production_qty} шт). На складе пока нет."
    return "Данных о производстве пока нет."


def _append_expected_date_to_copy(client_copy_text: str, label: str) -> str:
    return f"{client_copy_text} Ожидаем полный комплект на складе к {label}."


def _format_client_copy(kp_id: int, progress: SgpProgress, in_production_qty: int) -> str:
    n, m = progress.n, progress.m
    if n > 0 and in_production_qty > 0:
        return (
            f"Здравствуйте! По вашему заказу №{kp_id}: {n} из {m} шт уже на складе, "
            f"остальные ({in_production_qty} шт) в производстве. Можно забрать {n} шт."
        )
    if n > 0 and in_production_qty == 0:
        return (
            f"Здравствуйте! По вашему заказу №{kp_id}: {n} из {m} шт на складе, можно забрать."
        )
    if n == 0 and in_production_qty > 0:
        return (
            f"Здравствуйте! По вашему заказу №{kp_id}: заказ в производстве "
            f"({in_production_qty} шт), на складе пока нет."
        )
    return (
        f"Здравствуйте! По вашему заказу №{kp_id}: уточняем статус производства, "
        f"скоро сообщим."
    )


def _build_steps(
    *,
    progress: SgpProgress,
    completion_pct: float,
    in_plan_total: int,
    remaining_total: int,
) -> list[KpReadinessStep]:
    n, m = progress.n, progress.m
    pipeline_qty = in_plan_total + remaining_total

    if pipeline_qty > 0:
        production_state = KpReadinessStepState.ACTIVE
    elif n > 0:
        production_state = KpReadinessStepState.DONE
    else:
        production_state = KpReadinessStepState.PENDING

    if m > 0 and n == m:
        sgp_state = KpReadinessStepState.DONE
    elif n > 0:
        sgp_state = KpReadinessStepState.ACTIVE
    else:
        sgp_state = KpReadinessStepState.PENDING

    production_hint = f"{completion_pct:.0f}%" if completion_pct > 0 else None
    sgp_hint = f"{n}/{m}" if m > 0 else None

    step_defs: list[tuple[str, KpReadinessStepState, str | None]] = [
        ("kp", KpReadinessStepState.DONE, None),
        ("production", production_state, production_hint),
        ("sgp", sgp_state, sgp_hint),
        ("release", KpReadinessStepState.DISABLED, None),
        ("closed", KpReadinessStepState.DISABLED, None),
    ]
    return [
        KpReadinessStep(
            id=step_id,  # type: ignore[arg-type]
            label=_STEP_LABELS[step_id],
            state=state,
            hint=hint,
        )
        for step_id, state, hint in step_defs
    ]
