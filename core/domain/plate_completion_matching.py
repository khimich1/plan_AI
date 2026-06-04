"""Domain matching strategies for kp_plates row lookup (A2/A6)."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from typing import Callable, Sequence

LENGTH_TOLERANCE_M = 0.02
WIDTH_TOLERANCE_M = 0.05
WIDTH_TOLERANCE_STRICT_M = 0.01
STATUS_FILTER = "status IN ('в плане', 'в производстве')"
SELECT_COLS = "id, kp_id, plate_name, width_m, qty, nomenclature_id"

# Operational policy for allow_cross_kp=True:
# ai_docs/develop/guides/allow-cross-kp-runbook.md


@dataclass(frozen=True)
class PlateMatchContext:
    cur: sqlite3.Cursor
    plate_name: str
    length_m: float
    width_m: float
    load_class: int
    prefer_kp_id: int
    length_dm_raw: str | None
    allow_cross_kp: bool
    plan_ids: Sequence[str] | None
    width_clause: str
    width_params: tuple[float, ...]


def _normalize_plate_name(name: str) -> str:
    from core import plate_name as _pn

    return _pn.canonical(name)


def _width_parts(width_m: float) -> tuple[str, tuple[float, ...]]:
    if width_m and width_m > 0:
        return f"AND ABS(width_m - ?) < {WIDTH_TOLERANCE_M}", (width_m,)
    return "", ()


def _match_step_length_dm_raw(ctx: PlateMatchContext) -> tuple | None:
    if not ctx.length_dm_raw or not str(ctx.length_dm_raw).strip():
        return None
    ldr = str(ctx.length_dm_raw).strip()
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE kp_id = ? AND length_dm_raw = ? AND qty > 0
              AND {STATUS_FILTER} {ctx.width_clause}
            LIMIT 1
        """.replace("  ", " "),
        (ctx.prefer_kp_id, ldr, *ctx.width_params),
    )
    return ctx.cur.fetchone()


def _match_step_exact_name(ctx: PlateMatchContext) -> tuple | None:
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE kp_id = ? AND plate_name = ? AND qty > 0 AND {STATUS_FILTER}
            LIMIT 1
        """,
        (ctx.prefer_kp_id, ctx.plate_name),
    )
    return ctx.cur.fetchone()


def _match_step_canonical_scan(ctx: PlateMatchContext) -> tuple | None:
    from core import plate_name as _pn

    canon = _pn.canonical(ctx.plate_name)
    if not canon:
        return None
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE kp_id = ? AND qty > 0 AND {STATUS_FILTER}
        """,
        (ctx.prefer_kp_id,),
    )
    for cand in ctx.cur.fetchall():
        if _pn.canonical(cand[2]) == canon:
            return cand
    return None


def _match_step_normalized_name(ctx: PlateMatchContext) -> tuple | None:
    normalized_name = _normalize_plate_name(ctx.plate_name)
    if not normalized_name or normalized_name == ctx.plate_name:
        return None
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE kp_id = ? AND plate_name = ? AND qty > 0 AND {STATUS_FILTER}
            LIMIT 1
        """,
        (ctx.prefer_kp_id, normalized_name),
    )
    return ctx.cur.fetchone()


def _equiv_names(plate_name: str) -> list[str]:
    pairs = (
        ("59,9-12", "59,8-12"),
        ("59,8-12", "59,9-12"),
        ("61,1-12", "61,2-12"),
        ("61,2-12", "61,1-12"),
    )
    result: list[str] = []
    for old, new in pairs:
        if plate_name and old in plate_name:
            eq = plate_name.replace(old, new)
            if eq and eq != plate_name:
                result.append(eq)
    return result


def _match_step_equivalent_names(ctx: PlateMatchContext) -> tuple | None:
    for eq in _equiv_names(ctx.plate_name):
        ctx.cur.execute(
            f"""
                SELECT {SELECT_COLS} FROM kp_plates
                WHERE kp_id = ? AND plate_name = ? AND qty > 0 AND {STATUS_FILTER}
                LIMIT 1
            """,
            (ctx.prefer_kp_id, eq),
        )
        row = ctx.cur.fetchone()
        if row:
            return row
    return None


def _match_step_length_load(ctx: PlateMatchContext) -> tuple | None:
    if not ctx.length_m:
        return None
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE kp_id = ? AND {STATUS_FILTER} AND qty > 0
              AND ABS(length_m - ?) < ? {ctx.width_clause} AND load_class = ?
            LIMIT 1
        """,
        (ctx.prefer_kp_id, ctx.length_m, LENGTH_TOLERANCE_M, *ctx.width_params, ctx.load_class),
    )
    return ctx.cur.fetchone()


def _match_step_length_width_load(ctx: PlateMatchContext) -> tuple | None:
    if not (ctx.length_m and ctx.width_m):
        return None
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE kp_id = ? AND {STATUS_FILTER} AND qty > 0
              AND ABS(length_m - ?) < {LENGTH_TOLERANCE_M}
              AND ABS(width_m - ?) < {WIDTH_TOLERANCE_STRICT_M}
              AND load_class = ?
            LIMIT 1
        """,
        (ctx.prefer_kp_id, ctx.length_m, ctx.width_m, ctx.load_class),
    )
    return ctx.cur.fetchone()


def _match_step_length_load_relaxed_width(ctx: PlateMatchContext) -> tuple | None:
    if not ctx.length_m:
        return None
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE kp_id = ? AND {STATUS_FILTER} AND qty > 0
              AND ABS(length_m - ?) < {LENGTH_TOLERANCE_M}
              {ctx.width_clause} AND load_class = ?
            LIMIT 1
        """,
        (ctx.prefer_kp_id, ctx.length_m, *ctx.width_params, ctx.load_class),
    )
    return ctx.cur.fetchone()


def _match_step_cross_kp_by_plan(ctx: PlateMatchContext) -> tuple | None:
    if not (ctx.allow_cross_kp and ctx.length_m and ctx.plan_ids):
        return None
    placeholders = ",".join("?" * len(ctx.plan_ids))
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE plan_id IN ({placeholders}) AND {STATUS_FILTER} AND qty > 0
              AND ABS(length_m - ?) < {LENGTH_TOLERANCE_M}
              {ctx.width_clause} AND load_class = ?
            ORDER BY CASE WHEN kp_id = ? THEN 0 ELSE 1 END, id
            LIMIT 1
        """,
        (*ctx.plan_ids, ctx.length_m, *ctx.width_params, ctx.load_class, ctx.prefer_kp_id),
    )
    return ctx.cur.fetchone()


def _match_step_cross_kp_global(ctx: PlateMatchContext) -> tuple | None:
    if not (ctx.allow_cross_kp and ctx.length_m):
        return None
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE {STATUS_FILTER} AND qty > 0
              AND ABS(length_m - ?) < {LENGTH_TOLERANCE_M}
              {ctx.width_clause} AND load_class = ?
            ORDER BY CASE WHEN kp_id = ? THEN 0 ELSE 1 END, id
            LIMIT 1
        """,
        (ctx.length_m, *ctx.width_params, ctx.load_class, ctx.prefer_kp_id),
    )
    return ctx.cur.fetchone()


def _match_step_same_kp_length_fallback(ctx: PlateMatchContext) -> tuple | None:
    if not ctx.length_m:
        return None
    ctx.cur.execute(
        f"""
            SELECT {SELECT_COLS} FROM kp_plates
            WHERE qty > 0 AND {STATUS_FILTER}
              AND ABS(length_m - ?) < {LENGTH_TOLERANCE_M}
              {ctx.width_clause} AND load_class = ? AND kp_id = ?
            ORDER BY id
            LIMIT 1
        """,
        (ctx.length_m, *ctx.width_params, ctx.load_class, ctx.prefer_kp_id),
    )
    return ctx.cur.fetchone()


_MATCH_STEPS: tuple[Callable[[PlateMatchContext], tuple | None], ...] = (
    _match_step_length_dm_raw,
    _match_step_exact_name,
    _match_step_canonical_scan,
    _match_step_normalized_name,
    _match_step_equivalent_names,
    _match_step_length_load,
    _match_step_length_width_load,
    _match_step_length_load_relaxed_width,
    _match_step_cross_kp_by_plan,
    _match_step_cross_kp_global,
    _match_step_same_kp_length_fallback,
)


def find_kp_plate_row(
    cur: sqlite3.Cursor,
    plate_name: str,
    length_m: float,
    width_m: float,
    load_class: int,
    prefer_kp_id: int,
    *,
    length_dm_raw: str | None = None,
    allow_cross_kp: bool = False,
    plan_ids: Sequence[str] | None = None,
) -> tuple | None:
    """
    Find one ``kp_plates`` row for write-off (steps 0–10).

    ``allow_cross_kp`` must follow the operational runbook:
    ``ai_docs/develop/guides/allow-cross-kp-runbook.md``.
    """
    width_clause, width_params = _width_parts(width_m)
    ctx = PlateMatchContext(
        cur=cur,
        plate_name=plate_name,
        length_m=length_m,
        width_m=width_m,
        load_class=load_class,
        prefer_kp_id=prefer_kp_id,
        length_dm_raw=length_dm_raw,
        allow_cross_kp=allow_cross_kp,
        plan_ids=plan_ids,
        width_clause=width_clause,
        width_params=width_params,
    )
    for step in _MATCH_STEPS:
        row = step(ctx)
        if row:
            return row
    return None
