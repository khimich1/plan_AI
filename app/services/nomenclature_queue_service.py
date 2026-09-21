"""Очереди GUID/цен и резолв дублей (вкладка «Прайсы и 1С»)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from app.core.settings import get_settings
from app.schemas.nomenclature import (
    Create1cTaskOut,
    DuplicateCandidateOut,
    DuplicateResolveResponse,
    DuplicateTaskOut,
    GuidTasksResponse,
    PriceQueueResolveResponse,
    PriceTaskOut,
)
from core.duplicate_candidates import clear_candidates, list_candidates
from core.guid_demand import FIELD_DUPLICATE, list_open, sweep_guid_demand
from core.guid_queue import DuplicateTask, collect_guid_queue
from core.nomenclature_guid import get_guid_for_invoice, set_manual_choice, upsert
from core.pile_price_db import strip_trailing_u_suffix
from core.plate_guid_choice import normalize_plate_name, set_choice
from core.price_import_queue import apply_unit_price
from core.price_queue_db import list_open as list_open_price
from core.price_queue_db import set_resolved


class NomenclatureQueueError(ValueError):
    def __init__(self, message: str, *, status_code: int = 400) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.message = message


class NomenclatureQueueService:
    def __init__(self, db_path: Path | str | None = None) -> None:
        self._db_path = Path(db_path) if db_path is not None else None

    def list_tasks(self) -> GuidTasksResponse:
        conn = sqlite3.connect(str(self._pb_path()))
        try:
            snap = collect_guid_queue(conn)
            demand = list_open(conn)
            conn.commit()
        finally:
            conn.close()
        return GuidTasksResponse(
            to_create_1c=[
                Create1cTaskOut(
                    product_kind=row.product_kind,
                    mark=row.mark,
                    hint=row.hint,
                    field=row.field,
                    kp_ids=list(row.kp_ids),
                )
                for row in demand
                if row.field != FIELD_DUPLICATE
            ],
            to_price=[
                PriceTaskOut(
                    guid=item.guid,
                    product_kind=item.product_kind,
                    mark=item.mark,
                    name=item.name,
                )
                for item in snap.to_price
            ],
            duplicates=_duplicates_from_demand(demand, snap.duplicates),
        )

    def resolve_price(
        self,
        guid: str,
        price: float,
        *,
        resolved_by: str,
    ) -> PriceQueueResolveResponse:
        conn = sqlite3.connect(str(self._pb_path()))
        try:
            open_items = {item.guid: item for item in list_open_price(conn)}
            row = open_items.get(str(guid or "").strip().lower())
            if row is None:
                existing = conn.execute(
                    "SELECT guid, mark_norm, kind, name, state FROM price_queue WHERE guid = ?",
                    (str(guid or "").strip().lower(),),
                ).fetchone()
                if existing is None:
                    raise NomenclatureQueueError("Задача цены не найдена", status_code=404)
                if existing[4] == "resolved":
                    return PriceQueueResolveResponse(
                        guid=existing[0],
                        mark=existing[1],
                        product_kind=existing[2],
                        price=price,
                        state="resolved",
                    )
                raise NomenclatureQueueError("Задача цены не найдена", status_code=404)
        finally:
            conn.close()

        try:
            apply_unit_price(self._pb_path(), row.kind, row.mark_norm, price)
        except ValueError as exc:
            raise NomenclatureQueueError(str(exc), status_code=400) from exc

        conn = sqlite3.connect(str(self._pb_path()))
        try:
            upsert(
                conn,
                row.kind,
                row.mark_norm,
                guid_1c=row.guid,
                match_status="auto",
            )
            set_resolved(
                conn,
                row.guid,
                resolved_by=resolved_by,
                resolved_price=float(price),
            )
            conn.commit()
        finally:
            conn.close()
        return PriceQueueResolveResponse(
            guid=row.guid,
            mark=row.mark_norm,
            product_kind=row.kind,
            price=float(price),
            state="resolved",
        )

    def resolve_duplicate(
        self,
        *,
        scope: str,
        key: str,
        chosen_guid: str,
        note: str,
        decided_by: str,
        product_kind: str | None = None,
    ) -> DuplicateResolveResponse:
        scope_key = str(scope or "").strip().lower()
        guid = str(chosen_guid or "").strip().lower()
        conn = sqlite3.connect(str(self._pb_path()))
        try:
            if scope_key == "plate":
                if not _plate_guid_exists(conn, key, guid):
                    raise NomenclatureQueueError(
                        "кандидат исчез, задача открыта",
                        status_code=409,
                    )
                set_choice(
                    conn,
                    key,
                    chosen_guid=guid,
                    chosen_name=key,
                    decided_by=decided_by,
                    note=note,
                )
                conn.commit()
                sweep_guid_demand(conn)
                conn.commit()
                return DuplicateResolveResponse(
                    scope="plate",
                    key=key,
                    chosen_guid=guid,
                )
            kind = (product_kind or "pile").strip().lower()
            cands = list_candidates(conn, kind=kind, mark=key)
            if not any(item.guid == guid for item in cands):
                raise NomenclatureQueueError(
                    "кандидат исчез, задача открыта",
                    status_code=409,
                )
            set_manual_choice(
                conn,
                kind,
                key,
                guid_1c=guid,
                decided_by=decided_by,
                note=note or "по решению бухгалтерии",
            )
            clear_candidates(conn, kind, key)
            conn.commit()
            visible = get_guid_for_invoice(conn, kind, key)
            if visible != guid:
                raise NomenclatureQueueError("не удалось сохранить выбор GUID", status_code=500)
            sweep_guid_demand(conn)
            conn.commit()
        finally:
            conn.close()
        return DuplicateResolveResponse(
            scope="nonplate",
            key=key,
            chosen_guid=guid,
            match_status="manual",
        )

    def _pb_path(self) -> Path:
        if self._db_path is not None:
            return self._db_path
        return Path(get_settings().pb_db_path)


def _duplicates_from_demand(
    demand,
    catalog_dups: tuple[DuplicateTask, ...],
) -> list[DuplicateTaskOut]:
    kp_by_pair: dict[tuple[str, str], list[int]] = {}
    for row in demand:
        if row.field != FIELD_DUPLICATE:
            continue
        for key in _match_keys(row.product_kind, row.mark):
            bucket = kp_by_pair.setdefault((row.product_kind, key), [])
            for kp_id in row.kp_ids:
                if kp_id not in bucket:
                    bucket.append(kp_id)
    out: list[DuplicateTaskOut] = []
    for item in catalog_dups:
        kp_ids: list[int] = []
        for key in _match_keys(item.product_kind, item.key):
            for kp_id in kp_by_pair.get((item.product_kind, key), ()):
                if kp_id not in kp_ids:
                    kp_ids.append(kp_id)
        if not kp_ids or not item.candidates:
            continue
        out.append(_dup_out(item, kp_ids=kp_ids))
    return out


def _match_keys(product_kind: str, mark: str) -> set[str]:
    keys = {mark}
    if product_kind == "pile":
        stripped = strip_trailing_u_suffix(mark)
        if stripped:
            keys.add(stripped)
    if product_kind == "plate":
        try:
            keys.add(normalize_plate_name(mark))
        except ValueError:
            pass
    return keys


def _dup_out(item, *, kp_ids: list[int] | None = None) -> DuplicateTaskOut:
    return DuplicateTaskOut(
        scope=item.scope,
        key=item.key,
        product_kind=item.product_kind,
        candidates=[
            DuplicateCandidateOut(guid=guid, name=name, price=price)
            for guid, name, price in item.candidates
        ],
        kp_ids=list(kp_ids or ()),
    )


def _plate_guid_exists(conn: sqlite3.Connection, plate_name: str, guid: str) -> bool:
    tables = {
        row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
    }
    if "prays_plity" not in tables:
        return False
    row = conn.execute(
        """
        SELECT 1 FROM prays_plity
        WHERE "Товар" = ? COLLATE NOCASE
          AND lower("Уникальный идентификатор (Номенклатура)") = ?
        LIMIT 1
        """,
        (plate_name, guid),
    ).fetchone()
    return row is not None
