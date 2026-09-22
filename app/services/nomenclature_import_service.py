"""Загрузка выгрузки 1С «Прайс-лист» в nomenclature_guid (pb.db)."""

from __future__ import annotations

import sqlite3
import tempfile
from pathlib import Path
from typing import Sequence, TypeVar

from app.core.settings import get_settings
from app.schemas.nomenclature import (
    IMPORT_1C_LIST_LIMIT,
    AmbiguousMatchOut,
    DisappearedMarkOut,
    GuidWriteOut,
    Import1cResponse,
    Unmatched1COut,
)
from core.guid_demand import sweep_guid_demand
from core.nomenclature_guid import PRODUCT_KINDS, ensure_schema
from core.nomenclature_sync import (
    AmbiguousMatch,
    DisappearedMark,
    GuidWrite,
    SyncReport,
    Unmatched1C,
    infer_product_kind,
    infer_product_kind_from_names,
    sync_pricelist,
)
from core.price_import_queue import extract_mark_for_kind
from core.price_queue_db import ensure_schema as ensure_price_queue_schema
from core.price_queue_db import upsert_unmatched
from core.pricelist_1c_parser import Pricelist1CFormatError, PricelistRow, parse_pricelist_xls
from core.universal_report_1c_parser import (
    UniversalReport,
    UniversalReport1CFormatError,
    parse_universal_report_xlsx,
)

_T = TypeVar("_T")

_PLATE_ALIASES = frozenset({"plate", "plates", "plity", "плита", "плиты"})
_KIND_LIST = "pile, bridge_pile, fbs, stair_flight, stair_step"
_MSG_NEED_KIND = (
    "Не удалось определить группу изделий по имени файла. "
    "Передайте параметр product_kind "
    f"({_KIND_LIST})."
)
_MSG_UNKNOWN_KIND = f"Неизвестная группа изделий. Допустимы: {_KIND_LIST}."
_MSG_PLATES = "Плиты не загружаются в nomenclature_guid."
_MSG_EMPTY = "Пустой файл."
_MSG_AMBIGUOUS_KIND = (
    "Не удалось однозначно определить группу изделий по маркам. "
    f"Передайте параметр product_kind ({_KIND_LIST})."
)


class NomenclatureImportError(ValueError):
    """Ошибка формата/параметров загрузки — отдаём клиенту как 422."""


class NomenclatureImportService:
    def __init__(
        self,
        db_path: Path | str | None = None,
        weight_db_path: Path | str | None = None,
    ) -> None:
        self._db_path = Path(db_path) if db_path is not None else None
        self._weight_db_path = Path(weight_db_path) if weight_db_path is not None else None

    def import_file(
        self,
        file_bytes: bytes,
        filename: str | None,
        product_kind: str | None = None,
    ) -> Import1cResponse:
        suffix = Path(_safe_upload_name(filename)).suffix.lower()
        if suffix == ".xlsx":
            return self.import_xlsx(file_bytes, filename, product_kind)
        return self.import_xls(file_bytes, filename, product_kind)

    def import_xls(
        self,
        file_bytes: bytes,
        filename: str | None,
        product_kind: str | None = None,
    ) -> Import1cResponse:
        if not file_bytes:
            raise NomenclatureImportError(_MSG_EMPTY)
        safe_name = _safe_upload_name(filename)
        kind = resolve_import_product_kind(safe_name, product_kind)
        with tempfile.TemporaryDirectory(prefix="nomenclature-1c-") as tmp:
            path = Path(tmp) / safe_name
            path.write_bytes(file_bytes)
            try:
                rows = parse_pricelist_xls(path)
            except Pricelist1CFormatError:
                raise
            except ValueError as exc:
                raise NomenclatureImportError(str(exc)) from exc
        return self._sync_and_commit(rows, kind, partial=False)

    def import_xlsx(
        self,
        file_bytes: bytes,
        filename: str | None,
        product_kind: str | None = None,
    ) -> Import1cResponse:
        if not file_bytes:
            raise NomenclatureImportError(_MSG_EMPTY)
        safe_name = _safe_upload_name(filename, default="upload.xlsx")
        with tempfile.TemporaryDirectory(prefix="nomenclature-1c-") as tmp:
            path = Path(tmp) / safe_name
            path.write_bytes(file_bytes)
            try:
                report = parse_universal_report_xlsx(path)
            except UniversalReport1CFormatError:
                raise
            except ValueError as exc:
                raise NomenclatureImportError(str(exc)) from exc
        rows = _pricelist_from_universal(report)
        kind = _resolve_xlsx_kind(safe_name, product_kind, [row.name for row in report.rows])
        weights_updated, _skipped = _apply_xlsx_weights(
            self._weight_catalog_path(),
            report.rows,
            source_file=report.source_file,
        )
        return self._sync_and_commit(
            rows,
            kind,
            partial=report.partial,
            weights_updated=weights_updated,
        )

    def _sync_and_commit(
        self,
        rows: Sequence[PricelistRow],
        kind: str,
        *,
        partial: bool = False,
        weights_updated: int = 0,
    ) -> Import1cResponse:
        conn = sqlite3.connect(str(self._pb_db_path()))
        try:
            ensure_schema(conn)
            try:
                report = sync_pricelist(
                    conn,
                    rows,
                    product_kind=kind,
                    partial=partial,
                )
            except ValueError as exc:
                raise NomenclatureImportError(_sync_error_message(exc)) from exc
            _fill_price_queue(conn, kind, report.unmatched_1c)
            conn.commit()
            sweep_guid_demand(conn)
            conn.commit()
        finally:
            conn.close()
        mode = "partial" if partial else "full"
        return _report_to_response(report, kind, mode=mode, weights_updated=weights_updated)

    def _pb_db_path(self) -> Path:
        if self._db_path is not None:
            return self._db_path
        return Path(get_settings().pb_db_path)

    def _weight_catalog_path(self) -> Path:
        if self._weight_db_path is not None:
            return self._weight_db_path
        if self._db_path is not None:
            return self._db_path
        return Path(get_settings().plita_db_path)


def resolve_import_product_kind(
    filename: str | None,
    product_kind: str | None,
) -> str:
    """Явный kind или та же эвристика, что в core.nomenclature_sync.infer_product_kind.

    Имя файла: сваи → pile, мостовые → bridge_pile, блоки → fbs.
    Плиты, составные и смешанный ЛМ без явного product_kind — ошибка.
    """
    explicit = (product_kind or "").strip()
    if explicit:
        kind = explicit.lower()
        if kind in _PLATE_ALIASES:
            raise NomenclatureImportError(_MSG_PLATES)
        if kind not in PRODUCT_KINDS:
            raise NomenclatureImportError(
                f"Неизвестная группа изделий «{explicit}». "
                f"Допустимы: {_KIND_LIST}."
            )
        return kind
    inferred = infer_product_kind(filename)
    if inferred in PRODUCT_KINDS:
        return inferred
    name = Path(filename).name if filename else ""
    suffix = f" «{name}»" if name else ""
    raise NomenclatureImportError(
        f"Не удалось определить группу изделий по имени файла{suffix}. "
        f"Передайте параметр product_kind ({_KIND_LIST})."
    )


def _safe_upload_name(filename: str | None, *, default: str = "upload.xls") -> str:
    raw = (filename or "").strip()
    base = Path(raw).name if raw else default
    if not base or base in {".", ".."}:
        return default
    return base


def _sync_error_message(exc: ValueError) -> str:
    text = str(exc)
    if "product_kind is required" in text:
        return _MSG_NEED_KIND
    if "Unknown product_kind" in text:
        return _MSG_UNKNOWN_KIND
    return text


def _fill_price_queue(
    conn: sqlite3.Connection,
    kind: str,
    unmatched: Sequence[Unmatched1C],
) -> None:
    if not unmatched:
        return
    ensure_price_queue_schema(conn)
    items: list[tuple[str, str, str, str]] = []
    for row in unmatched:
        mark = extract_mark_for_kind(kind, row.name) or row.name
        items.append((row.guid, mark, kind, row.name))
    upsert_unmatched(conn, items)


def _apply_xlsx_weights(
    db_path: Path,
    rows: Sequence[object],
    *,
    source_file: str,
) -> tuple[int, int]:
    from core.product_weight_catalog import apply_universal_report_weights

    return apply_universal_report_weights(
        str(db_path),
        rows,
        source_file=source_file,
    )


def _resolve_xlsx_kind(
    filename: str | None,
    product_kind: str | None,
    names: Sequence[str],
) -> str:
    explicit = (product_kind or "").strip()
    if explicit:
        return resolve_import_product_kind(filename, explicit)
    inferred_file = infer_product_kind(filename)
    if inferred_file in PRODUCT_KINDS:
        return inferred_file
    inferred_names = infer_product_kind_from_names(names)
    if inferred_names in _PLATE_ALIASES or inferred_names == "plate":
        raise NomenclatureImportError(_MSG_PLATES)
    if inferred_names in PRODUCT_KINDS:
        return inferred_names
    raise NomenclatureImportError(_MSG_AMBIGUOUS_KIND)


def _pricelist_from_universal(report: UniversalReport) -> list[PricelistRow]:
    return [
        PricelistRow(
            guid=row.guid,
            name=row.name,
            price=row.price,
            unit="",
            source_file=row.source_file,
            row_index=row.row_index,
        )
        for row in report.rows
    ]


def _report_to_response(
    report: SyncReport,
    kind: str,
    *,
    mode: str = "full",
    weights_updated: int = 0,
) -> Import1cResponse:
    new_n = len(report.new_guids)
    amb_n = len(report.ambiguous)
    dis_n = len(report.disappeared)
    summary = (
        f"+{new_n} GUID, {report.waiting_price} ждут цены, "
        f"{amb_n} неоднозначных, {dis_n} исчезли из 1С"
    )
    return Import1cResponse(
        product_kind=kind,
        summary=summary,
        new_guids_count=new_n,
        waiting_price=report.waiting_price,
        ambiguous_count=amb_n,
        disappeared_count=dis_n,
        unmatched_1c_count=len(report.unmatched_1c),
        unchanged=report.unchanged,
        updated_guids_count=len(report.updated_guids),
        mode=mode,
        weights_updated=weights_updated,
        list_limit=IMPORT_1C_LIST_LIMIT,
        new_guids=[_guid_out(item) for item in _cap(report.new_guids)],
        updated_guids=[_guid_out(item) for item in _cap(report.updated_guids)],
        ambiguous=[_ambiguous_out(item) for item in _cap(report.ambiguous)],
        disappeared=[_disappeared_out(item) for item in _cap(report.disappeared)],
        unmatched_1c=[_unmatched_out(item) for item in _cap(report.unmatched_1c)],
    )


def _cap(items: Sequence[_T], limit: int = IMPORT_1C_LIST_LIMIT) -> list[_T]:
    return list(items[:limit])


def _guid_out(item: GuidWrite) -> GuidWriteOut:
    return GuidWriteOut(
        product_kind=item.product_kind,
        mark=item.mark,
        field=item.field,
        guid=item.guid,
    )


def _ambiguous_out(item: AmbiguousMatch) -> AmbiguousMatchOut:
    return AmbiguousMatchOut(
        product_kind=item.product_kind,
        mark=item.mark,
        name=item.name,
        guids=list(item.guids),
        note=item.note,
    )


def _disappeared_out(item: DisappearedMark) -> DisappearedMarkOut:
    return DisappearedMarkOut(
        product_kind=item.product_kind,
        mark=item.mark,
        guid_1c=item.guid_1c,
        match_status=item.match_status,
    )


def _unmatched_out(item: Unmatched1C) -> Unmatched1COut:
    return Unmatched1COut(
        name=item.name,
        guid=item.guid,
        row_index=item.row_index,
        source_file=item.source_file,
    )


__all__ = [
    "NomenclatureImportError",
    "NomenclatureImportService",
    "resolve_import_product_kind",
]
