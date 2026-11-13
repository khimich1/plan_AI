#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сверка очереди guid_demand (pb.db) с составом КП (plita.db).

Источник правды — КП. Строка очереди производная: её список КП и состояние
приводит эта сверка (спека pochemu-net-guid, D-reconcile). Функция идемпотентна,
сама не коммитит и не пишет в plita.db.
"""

from __future__ import annotations

import sqlite3
from collections import defaultdict
from dataclasses import dataclass
from typing import Optional

from core.guid_demand import (
    GuidDemandRow,
    ensure_schema,
    list_open,
    record_guid_demand,
    set_kp_ids,
    set_resolved,
)
from core.guid_gate import OrderLine, check_invoice_guids

ACTION_CLOSED = "closed"
ACTION_NARROWED = "narrowed"
ACTION_REBOUND = "rebound"
ACTION_OPENED = "opened"

_ACTION_LABELS = {
    ACTION_CLOSED: "закрыта",
    ACTION_NARROWED: "сужена",
    ACTION_REBOUND: "перепривязана",
    ACTION_OPENED: "открыта новая",
}

# (product_kind очереди, таблица КП, столбец с маркой)
_KP_SOURCES = (
    ("plate", "kp_plates", "plate_name"),
    ("pile", "kp_piles", "mark"),
    ("stair_step", "kp_steps", "mark"),
    ("stair_flight", "kp_marches", "mark"),
    ("bridge_pile", "kp_bridge_piles", "mark"),
    ("composite_pile", "kp_composite_piles", "mark"),
    ("fbs", "kp_fbs", "mark"),
)

_Key = tuple[str, str]


@dataclass(frozen=True)
class ReconcileChange:
    product_kind: str
    mark: str
    field: str
    action: str
    kp_ids_before: tuple[int, ...]
    kp_ids_after: tuple[int, ...]


@dataclass(frozen=True)
class ReconcileReport:
    changes: tuple[ReconcileChange, ...]
    unchanged: int

    @property
    def has_changes(self) -> bool:
        return bool(self.changes)


def _collect_kp_index(
    plita_conn: sqlite3.Connection,
) -> tuple[dict[_Key, list[int]], dict[int, list[OrderLine]]]:
    """Что сейчас лежит в КП: (вид, марка) → КП и КП → строки для резолвера."""
    kps_by_mark: dict[_Key, set[int]] = defaultdict(set)
    lines_by_kp: dict[int, dict[_Key, OrderLine]] = defaultdict(dict)
    for kind, table, column in _KP_SOURCES:
        cur = plita_conn.execute(
            f"SELECT t.kp_id, t.{column} FROM {table} t "
            "JOIN KP_offers o ON o.kp_id = t.kp_id"
        )
        for kp_id, raw_mark in cur:
            mark = str(raw_mark or "").strip()
            if not mark:
                continue
            kps_by_mark[(kind, mark)].add(int(kp_id))
            lines_by_kp[int(kp_id)][(kind, mark)] = OrderLine(kind, mark)
    index = {key: sorted(ids) for key, ids in kps_by_mark.items()}
    return index, {kp_id: list(lines.values()) for kp_id, lines in lines_by_kp.items()}


def _classify(before: tuple[int, ...], after: tuple[int, ...]) -> Optional[str]:
    if set(before) == set(after):
        return None
    if not after:
        return ACTION_CLOSED
    if set(after) < set(before):
        return ACTION_NARROWED
    return ACTION_REBOUND


def _sync_row(
    pb_conn: sqlite3.Connection,
    row: GuidDemandRow,
    actual: list[int],
) -> Optional[ReconcileChange]:
    action = _classify(row.kp_ids, tuple(actual))
    if action is None:
        return None
    if action == ACTION_CLOSED:
        set_kp_ids(pb_conn, row, [])
        set_resolved(pb_conn, row)
    else:
        set_kp_ids(pb_conn, row, actual)
    return ReconcileChange(
        product_kind=row.product_kind,
        mark=row.mark,
        field=row.field,
        action=action,
        kp_ids_before=row.kp_ids,
        kp_ids_after=tuple(actual),
    )


def _recheck_dropped_kps(
    pb_conn: sqlite3.Connection,
    dropped: set[int],
    lines_by_kp: dict[int, list[OrderLine]],
) -> None:
    """Выпавшие КП: резолвер по их текущему составу, новые дыры → в очередь."""
    for kp_id in sorted(dropped):
        lines = lines_by_kp.get(kp_id)
        if not lines:
            continue
        report = check_invoice_guids(lines, pb_conn)
        if report.missing:
            record_guid_demand(pb_conn, kp_id, report.missing)


def _opened_changes(
    pb_conn: sqlite3.Connection,
    open_before: set[tuple[str, str, str]],
    index: dict[_Key, list[int]],
) -> list[ReconcileChange]:
    changes: list[ReconcileChange] = []
    for row in list_open(pb_conn):
        key = (row.product_kind, row.mark, row.field)
        if key in open_before:
            continue
        actual = index.get((row.product_kind, row.mark), [])
        if tuple(actual) != row.kp_ids:
            set_kp_ids(pb_conn, row, actual)
        changes.append(
            ReconcileChange(
                product_kind=row.product_kind,
                mark=row.mark,
                field=row.field,
                action=ACTION_OPENED,
                kp_ids_before=(),
                kp_ids_after=tuple(actual),
            )
        )
    return changes


def reconcile_guid_demand(
    pb_conn: sqlite3.Connection,
    plita_conn: sqlite3.Connection,
) -> ReconcileReport:
    """Привести открытые строки guid_demand к составу КП. Не коммитит."""
    ensure_schema(pb_conn)
    index, lines_by_kp = _collect_kp_index(plita_conn)
    rows = list_open(pb_conn)
    open_before = {(r.product_kind, r.mark, r.field) for r in rows}

    changes: list[ReconcileChange] = []
    dropped: set[int] = set()
    for row in rows:
        actual = index.get((row.product_kind, row.mark), [])
        change = _sync_row(pb_conn, row, actual)
        if change is None:
            continue
        changes.append(change)
        dropped.update(set(row.kp_ids) - set(actual))

    if dropped:
        _recheck_dropped_kps(pb_conn, dropped, lines_by_kp)
        changes.extend(_opened_changes(pb_conn, open_before, index))
    return ReconcileReport(
        changes=tuple(changes),
        unchanged=len(rows) - sum(1 for c in changes if c.action != ACTION_OPENED),
    )


def _ids_text(kp_ids: tuple[int, ...]) -> str:
    return ", ".join(str(kp_id) for kp_id in kp_ids) if kp_ids else "нет"


def _change_line(change: ReconcileChange) -> str:
    label = _ACTION_LABELS[change.action]
    if change.action == ACTION_NARROWED:
        label = f"{label} до КП {_ids_text(change.kp_ids_after)}"
    return (
        f"- {label}: {change.product_kind} «{change.mark}» "
        f"(КП: {_ids_text(change.kp_ids_before)} → {_ids_text(change.kp_ids_after)})"
    )


def format_report(
    report: ReconcileReport,
    *,
    apply: bool,
    pb_path: object = None,
    plita_path: object = None,
) -> str:
    """Отчёт сверки по-русски: по строке «было → стало», сводка, «не изменилось»."""
    title = "Сверка очереди «Завести в 1С» с КП — "
    title += "записано" if apply else "сухой прогон, ничего не записано"
    lines = [title]
    if pb_path is not None or plita_path is not None:
        lines.append(f"База очереди (pb.db): {pb_path}")
        lines.append(f"База КП (plita.db): {plita_path}")
    if not report.has_changes:
        lines.append(f"Изменений нет. Не изменилось: {report.unchanged}.")
        return "\n".join(lines)
    lines.extend(_change_line(change) for change in report.changes)
    counts = {action: 0 for action in _ACTION_LABELS}
    for change in report.changes:
        counts[change.action] += 1
    lines.append(
        f"Итого: закрыто {counts[ACTION_CLOSED]}, сужено {counts[ACTION_NARROWED]}, "
        f"перепривязано {counts[ACTION_REBOUND]}, открыто новых {counts[ACTION_OPENED]}. "
        f"Не изменилось: {report.unchanged}."
    )
    return "\n".join(lines)


__all__ = [
    "ACTION_CLOSED",
    "ACTION_NARROWED",
    "ACTION_OPENED",
    "ACTION_REBOUND",
    "ReconcileChange",
    "ReconcileReport",
    "format_report",
    "reconcile_guid_demand",
]
