#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Живые очереди GUID: 🏭 завести в 1С, 💰 ввести цену, ⚠️ дубли."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from typing import Optional

from core.duplicate_candidates import ensure_schema as ensure_dup_schema
from core.duplicate_candidates import list_candidates
from core.nomenclature_guid import MATCH_STATUSES, ensure_schema as ensure_guid_schema
from core.nomenclature_guid import iter_by_state
from core.plate_guid_choice import ensure_schema as ensure_choice_schema
from core.plate_guid_choice import get_choice
from core.price_queue_db import ensure_schema as ensure_price_schema
from core.price_queue_db import list_open


@dataclass(frozen=True)
class Create1cTask:
    product_kind: str
    mark: str
    hint: str
    field: str = "guid_1c"


@dataclass(frozen=True)
class PriceTask:
    guid: str
    product_kind: str
    mark: str
    name: str


@dataclass(frozen=True)
class DuplicateTask:
    scope: str
    key: str
    product_kind: str
    candidates: tuple[tuple[str, str, Optional[float]], ...]


@dataclass(frozen=True)
class GuidQueueSnapshot:
    to_create_1c: tuple[Create1cTask, ...]
    to_price: tuple[PriceTask, ...]
    duplicates: tuple[DuplicateTask, ...]


def collect_guid_queue(conn: sqlite3.Connection) -> GuidQueueSnapshot:
    """Собрать живые очереди из pb.db (nomenclature_guid / price_queue / дубли)."""
    ensure_guid_schema(conn)
    ensure_choice_schema(conn)
    ensure_dup_schema(conn)
    ensure_price_schema(conn)
    return GuidQueueSnapshot(
        to_create_1c=_collect_create(conn),
        to_price=_collect_price(conn),
        duplicates=_collect_duplicates(conn),
    )


def _collect_create(conn: sqlite3.Connection) -> tuple[Create1cTask, ...]:
    tasks: list[Create1cTask] = []
    seen: set[tuple[str, str, str]] = set()
    for status in MATCH_STATUSES:
        for row in iter_by_state(conn, status):
            if not row.guid_1c:
                item = Create1cTask(
                    product_kind=row.product_kind,
                    mark=row.mark,
                    hint="заведите карточку в 1С и загрузите отчёт",
                    field="guid_1c",
                )
                key = (item.product_kind, item.mark, item.field)
                if key not in seen:
                    seen.add(key)
                    tasks.append(item)
            elif row.product_kind == "pile" and not row.guid_1c_u:
                item = Create1cTask(
                    product_kind=row.product_kind,
                    mark=row.mark,
                    hint="заведите карточку «у» в 1С и загрузите отчёт",
                    field="guid_1c_u",
                )
                key = (item.product_kind, item.mark, item.field)
                if key not in seen:
                    seen.add(key)
                    tasks.append(item)
    tasks.sort(key=lambda item: (item.product_kind, item.mark, item.field))
    return tuple(tasks)


def _collect_price(conn: sqlite3.Connection) -> tuple[PriceTask, ...]:
    return tuple(
        PriceTask(
            guid=item.guid,
            product_kind=item.kind,
            mark=item.mark_norm,
            name=item.name,
        )
        for item in list_open(conn)
    )


def _collect_duplicates(conn: sqlite3.Connection) -> tuple[DuplicateTask, ...]:
    grouped: dict[tuple[str, str], list[tuple[str, str, Optional[float]]]] = {}
    for item in list_candidates(conn):
        key = (item.kind, item.mark_norm)
        grouped.setdefault(key, []).append((item.guid, item.name, item.price))
    tasks: list[DuplicateTask] = []
    for (kind, mark), cands in sorted(grouped.items()):
        tasks.append(
            DuplicateTask(
                scope="nonplate",
                key=mark,
                product_kind=kind,
                candidates=tuple(cands),
            )
        )
    tasks.extend(_plate_duplicate_tasks(conn))
    return tuple(tasks)


def _plate_duplicate_tasks(conn: sqlite3.Connection) -> list[DuplicateTask]:
    tables = {
        row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    if "prays_plity" not in tables:
        return []
    cur = conn.execute(
        """
        SELECT "Товар",
               GROUP_CONCAT("Уникальный идентификатор (Номенклатура)", char(31))
        FROM prays_plity
        GROUP BY "Товар" COLLATE NOCASE
        HAVING COUNT(DISTINCT "Уникальный идентификатор (Номенклатура)") > 1
        """
    )
    tasks: list[DuplicateTask] = []
    for name, packed in cur:
        if get_choice(conn, str(name or "")) is not None:
            continue
        guids = []
        seen: set[str] = set()
        for guid in str(packed or "").split("\x1f"):
            key = guid.strip().lower()
            if key and key not in seen:
                seen.add(key)
                guids.append((key, str(name), None))
        if len(guids) > 1:
            tasks.append(
                DuplicateTask(
                    scope="plate",
                    key=str(name),
                    product_kind="plate",
                    candidates=tuple(guids),
                )
            )
    return tasks


__all__ = [
    "Create1cTask",
    "DuplicateTask",
    "GuidQueueSnapshot",
    "PriceTask",
    "collect_guid_queue",
]
