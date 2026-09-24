#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сверка выгрузки 1С «Прайс-лист» с таблицей nomenclature_guid.

Строки выгрузки сопоставляются с марками справочника. Если строки ещё
нет, единственный GUID пишется на марку из прайса этой группы или из
открытого guid_demand (guid_1c / guid_1c_u). Прайс не меняется. Имя вне
прайса и вне спроса остаётся unmatched_1c. Несколько GUID или несколько
марок на одно имя — без записи GUID. Плиты в эту таблицу не пишутся.

Правило смены GUID: unique match при status=auto может обновить GUID,
если в файле ровно один GUID на нормализованное имя. status=manual
никогда не перезаписывается без overwrite_manual=True. Автовыбора
среди дублей нет.
"""

from __future__ import annotations

import re
import sqlite3
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional, Sequence

from core.nomenclature_guid import (
    MATCH_STATUSES,
    PRODUCT_KINDS,
    NomenclatureGuidRow,
    ensure_schema,
    get_by_mark,
    iter_by_state,
    upsert,
)
from core.pile_line_parser import is_reinforced_pile_mark
from core.pile_price_db import strip_trailing_u_suffix
from core.pricelist_1c_parser import PricelistRow

FIELD_PLAIN = "guid_1c"
FIELD_U = "guid_1c_u"

_PRICE_TABLE_BY_KIND = {
    "pile": "pile_prices",
    "bridge_pile": "bridge_pile_prices",
    "fbs": "fbs_prices",
    "stair_flight": "march_prices",
    "stair_step": "step_prices",
}
_DEMAND_OPEN = "open"

_PLATE_ALIASES = frozenset({"plate", "plates", "plity", "плита", "плиты"})

# Как в _tmp_guid_match.py / scripts/build_guid_queue.py: латиница ≡ кириллица.
_LAT2CYR = str.maketrans(
    {
        "c": "с",
        "C": "С",
        "t": "т",
        "T": "Т",
        "b": "в",
        "B": "В",
        "p": "р",
        "P": "Р",
        "a": "а",
        "A": "А",
        "e": "е",
        "E": "Е",
        "h": "н",
        "H": "Н",
        "m": "м",
        "M": "М",
        "k": "к",
        "K": "К",
        "x": "х",
        "X": "Х",
        "o": "о",
        "O": "О",
        "y": "у",
        "Y": "У",
    }
)

_BRIDGE_MARK_RE = re.compile(
    r"^[CcСс]\s*(\d+)[-\.](\d+)\s*([TТtтBВbв])\s*(\d+)$"
)
_STEP_MARK_RE = re.compile(r"^ЛС\s*(\d+)(.*)$", re.I)
_PILE_LEADING_C_RE = re.compile(r"^[CcСс]\s*")


@dataclass(frozen=True, slots=True)
class GuidWrite:
    product_kind: str
    mark: str
    field: str
    guid: str


@dataclass(frozen=True, slots=True)
class AmbiguousMatch:
    product_kind: str
    mark: Optional[str]
    name: str
    guids: tuple[str, ...]
    note: str


@dataclass(frozen=True, slots=True)
class DisappearedMark:
    product_kind: str
    mark: str
    guid_1c: Optional[str]
    match_status: str


@dataclass(frozen=True, slots=True)
class Unmatched1C:
    name: str
    guid: str
    row_index: int
    source_file: str


@dataclass(frozen=True, slots=True)
class SyncReport:
    """Отчёт сверки: +N GUID, M ждут цены, K неоднозначных, J исчезли из 1С."""

    new_guids: tuple[GuidWrite, ...] = ()
    waiting_price: int = 0
    ambiguous: tuple[AmbiguousMatch, ...] = ()
    disappeared: tuple[DisappearedMark, ...] = ()
    unchanged: int = 0
    unmatched_1c: tuple[Unmatched1C, ...] = ()
    updated_guids: tuple[GuidWrite, ...] = ()


@dataclass
class _MarkPlan:
    ambiguous_guids: list[str] = field(default_factory=list)
    ambiguous_names: list[str] = field(default_factory=list)
    fills: dict[str, str] = field(default_factory=dict)
    fill_conflicts: list[str] = field(default_factory=list)


def normalize_name(value: str) -> str:
    """Lowercase, Latin/Cyrillic lookalikes, no spaces. C ≡ С."""
    text = str(value).strip().lower().replace("ё", "е").translate(_LAT2CYR)
    text = text.replace("–", "-").replace("—", "-").replace(",", ".").replace("*", "")
    return re.sub(r"\s+", "", text)


def infer_product_kind(source_file: str | None) -> str | None:
    """Один файл 1С = один kind. Плиты/составные/смешанный ЛМ → None."""
    if not source_file:
        return None
    name = Path(source_file).name.lower()
    if "плит" in name:
        return None
    if "мостов" in name:
        if "составн" in name:
            return None
        return "bridge_pile"
    if "составн" in name:
        return None
    if "сва" in name:
        return "pile"
    if "блок" in name:
        return "fbs"
    return None


_SKIP_NAME_RE = re.compile(
    r"площадк|перемычк|прогон|фундамент|составн",
    re.I | re.UNICODE,
)
_PILE_NAME_RE = re.compile(r"(?:^|\s)сваи\s+|^[cс]\s*\d", re.I | re.UNICODE)
_BRIDGE_NAME_RE = re.compile(r"мостов", re.I | re.UNICODE)
_FBS_NAME_RE = re.compile(r"фбс|блоки\s+фбс", re.I | re.UNICODE)
_MARCH_NAME_RE = re.compile(r"лестничные\s+марши|\d*лм\b", re.I | re.UNICODE)
_STEP_NAME_RE = re.compile(r"лестничные\s+ступени|\bлс[\s\-]?\d", re.I | re.UNICODE)
_PLATE_NAME_RE = re.compile(r"плит|\bп[бк]\s*\d", re.I | re.UNICODE)


def infer_product_kind_from_names(names: Sequence[str]) -> str | None:
    """Один kind по маркам; смешанный/пустой/плиты → None."""
    kinds: set[str] = set()
    for raw in names:
        name = str(raw or "").strip()
        if not name or _SKIP_NAME_RE.search(name):
            continue
        kind = _kind_from_nomenclature_name(name)
        if kind:
            kinds.add(kind)
    if len(kinds) == 1:
        return next(iter(kinds))
    return None


def _kind_from_nomenclature_name(name: str) -> str | None:
    if _BRIDGE_NAME_RE.search(name):
        return "bridge_pile"
    if _FBS_NAME_RE.search(name):
        return "fbs"
    if _MARCH_NAME_RE.search(name):
        return "stair_flight"
    if _STEP_NAME_RE.search(name):
        return "stair_step"
    if _PLATE_NAME_RE.search(name):
        return "plate"
    if _PILE_NAME_RE.search(name) or name.lower().startswith("сваи"):
        return "pile"
    return None


def sync_pricelist(
    conn: sqlite3.Connection,
    rows: Sequence[PricelistRow],
    product_kind: str | None = None,
    *,
    overwrite_manual: bool = False,
    partial: bool = False,
) -> SyncReport:
    """Сопоставить выгрузку с nomenclature_guid и применить обновления.

    Caller owns commit. Строки таблицы не удаляются. Плиты и неизвестный
    kind не вставляются. partial=True (отбор 1С) — не заполняет disappeared.
    Прайс-таблицы только читаются.
    """
    ensure_schema(conn)
    incoming = [row for row in rows if row.guid and row.name]
    kind = _coerce_kind(product_kind, incoming)
    if kind is None:
        return _skipped_report(incoming)

    db_rows = _rows_for_kind(conn, kind)
    index = _build_name_index(kind, db_rows)
    admission = _index_marks(kind, _admission_marks(conn, kind))
    grouped = _group_by_normalized_name(incoming)

    plans: dict[str, _MarkPlan] = defaultdict(_MarkPlan)
    unmatched: list[Unmatched1C] = []
    ambiguous_unmatched: list[AmbiguousMatch] = []
    matched_marks: set[str] = set()

    for norm_name, group in grouped.items():
        guids = tuple(sorted({item.guid.strip().lower() for item in group}))
        targets = index.get(norm_name) or admission.get(norm_name, [])
        sample = group[0]
        if len(guids) > 1:
            _record_ambiguous(
                kind,
                sample.name,
                guids,
                targets,
                plans,
                matched_marks,
                ambiguous_unmatched,
            )
            continue
        if not targets:
            unmatched.extend(_unmatched_from_group(group))
            continue
        selected = _select_target(kind, norm_name, targets)
        if selected is None:
            _record_ambiguous(
                kind,
                sample.name,
                guids,
                targets,
                plans,
                matched_marks,
                ambiguous_unmatched,
            )
            continue
        mark, field_name = selected
        matched_marks.add(mark)
        _add_fill(plans[mark], field_name, guids[0])

    new_guids: list[GuidWrite] = []
    updated_guids: list[GuidWrite] = []
    ambiguous: list[AmbiguousMatch] = list(ambiguous_unmatched)
    unchanged = 0

    rows_by_mark = {row.mark: row for row in db_rows}
    for mark, plan in plans.items():
        db_row = rows_by_mark.get(mark) or _placeholder_row(kind, mark)
        writes, amb, n_unchanged = _apply_mark(
            conn,
            db_row,
            plan,
            overwrite_manual=overwrite_manual,
        )
        for item in writes:
            if item[0] == "new":
                new_guids.append(item[1])
            else:
                updated_guids.append(item[1])
        if amb is not None:
            ambiguous.append(amb)
        unchanged += n_unchanged

    disappeared = ()
    if not partial:
        disappeared = tuple(
            DisappearedMark(
                product_kind=row.product_kind,
                mark=row.mark,
                guid_1c=row.guid_1c,
                match_status=row.match_status,
            )
            for row in db_rows
            if row.mark not in matched_marks
        )
    report = SyncReport(
        new_guids=tuple(new_guids),
        waiting_price=0,
        ambiguous=tuple(ambiguous),
        disappeared=disappeared,
        unchanged=unchanged,
        unmatched_1c=tuple(unmatched),
        updated_guids=tuple(updated_guids),
    )
    _persist_ambiguous_candidates(conn, kind, report.ambiguous, incoming)
    return report


def _coerce_kind(
    product_kind: str | None,
    rows: Sequence[PricelistRow],
) -> str | None:
    if product_kind is not None and str(product_kind).strip():
        kind = str(product_kind).strip().lower()
        if kind in _PLATE_ALIASES:
            return None
        if kind not in PRODUCT_KINDS:
            raise ValueError(f"Unknown product_kind: {product_kind!r}")
        return kind
    source = next((row.source_file for row in rows if row.source_file), None)
    inferred = infer_product_kind(source)
    if inferred in PRODUCT_KINDS:
        return inferred
    if _is_skip_source(source):
        return None
    raise ValueError(
        "product_kind is required (one 1C file = one kind; "
        "plates are not written to nomenclature_guid)"
    )


def _is_skip_source(source_file: str | None) -> bool:
    if not source_file:
        return False
    name = Path(source_file).name.lower()
    return "плит" in name or "составн" in name


def _skipped_report(rows: Sequence[PricelistRow]) -> SyncReport:
    return SyncReport(
        unmatched_1c=tuple(_unmatched_from_group(rows)),
    )


def _unmatched_from_group(group: Sequence[PricelistRow]) -> list[Unmatched1C]:
    return [
        Unmatched1C(
            name=row.name,
            guid=row.guid,
            row_index=row.row_index,
            source_file=row.source_file,
        )
        for row in group
    ]


def _rows_for_kind(
    conn: sqlite3.Connection,
    kind: str,
) -> list[NomenclatureGuidRow]:
    found: list[NomenclatureGuidRow] = []
    for status in MATCH_STATUSES:
        for row in iter_by_state(conn, status):
            if row.product_kind == kind:
                found.append(row)
    found.sort(key=lambda row: row.mark)
    return found


def _group_by_normalized_name(
    rows: Sequence[PricelistRow],
) -> dict[str, list[PricelistRow]]:
    grouped: dict[str, list[PricelistRow]] = defaultdict(list)
    for row in rows:
        grouped[normalize_name(row.name)].append(row)
    return grouped


def _build_name_index(
    kind: str,
    db_rows: Sequence[NomenclatureGuidRow],
) -> dict[str, list[tuple[str, str]]]:
    return _index_marks(kind, (row.mark for row in db_rows))


def _index_marks(
    kind: str,
    marks: Iterable[str],
) -> dict[str, list[tuple[str, str]]]:
    index: dict[str, list[tuple[str, str]]] = defaultdict(list)
    seen: dict[str, set[tuple[str, str]]] = defaultdict(set)
    for mark in marks:
        text = str(mark or "")
        if not text.strip():
            continue
        for candidate, field_name in _candidates_for_mark(kind, text):
            key = (text, field_name)
            if key in seen[candidate]:
                continue
            seen[candidate].add(key)
            index[candidate].append(key)
    return index


def _table_exists(conn: sqlite3.Connection, table: str) -> bool:
    found = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table,),
    ).fetchone()
    return found is not None


def _price_marks(conn: sqlite3.Connection, kind: str) -> list[str]:
    table = _PRICE_TABLE_BY_KIND.get(kind)
    if table is None or not _table_exists(conn, table):
        return []
    found: list[str] = []
    seen: set[str] = set()
    for (raw,) in conn.execute(f"SELECT DISTINCT mark FROM {table}"):
        text = str(raw or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        found.append(text)
    return found


def _demand_catalog_mark(kind: str, mark: str) -> str:
    """Марка справочника для спроса: у сваи «у» GUID живёт на марке без суффикса."""
    if kind != "pile" or not is_reinforced_pile_mark(mark):
        return mark
    stripped = strip_trailing_u_suffix(mark).strip()
    return stripped or mark


def _open_demand_marks(conn: sqlite3.Connection, kind: str) -> list[str]:
    if not _table_exists(conn, "guid_demand"):
        return []
    rows = conn.execute(
        """
        SELECT DISTINCT mark FROM guid_demand
        WHERE product_kind = ? AND state = ? AND field IN (?, ?)
        """,
        (kind, _DEMAND_OPEN, FIELD_PLAIN, FIELD_U),
    )
    found: list[str] = []
    seen: set[str] = set()
    for (raw,) in rows:
        text = _demand_catalog_mark(kind, str(raw or "").strip())
        if not text or text in seen:
            continue
        seen.add(text)
        found.append(text)
    return found


def _admission_marks(conn: sqlite3.Connection, kind: str) -> list[str]:
    marks = _price_marks(conn, kind)
    seen = set(marks)
    for mark in _open_demand_marks(conn, kind):
        if mark in seen:
            continue
        seen.add(mark)
        marks.append(mark)
    return marks


def _placeholder_row(kind: str, mark: str) -> NomenclatureGuidRow:
    """Пустая марка: _apply_mark создаст строку справочника, прайс не трогает."""
    return NomenclatureGuidRow(
        product_kind=kind,
        mark=mark,
        guid_1c=None,
        guid_1c_u=None,
        match_status="missing",
        match_note=None,
        updated_at="",
    )


def _candidates_for_mark(kind: str, mark: str) -> list[tuple[str, str]]:
    if kind == "pile":
        return _pile_candidates(mark)
    if kind == "bridge_pile":
        return _bridge_candidates(mark)
    if kind == "fbs":
        return _prefixed_candidates(mark, "Блоки")
    if kind == "stair_flight":
        return _prefixed_candidates(mark, "Лестничные марши")
    if kind == "stair_step":
        return _step_candidates(mark)
    return [(normalize_name(mark), FIELD_PLAIN)]


def _pile_base(mark: str) -> str:
    return _PILE_LEADING_C_RE.sub("С ", mark.strip().rstrip("* ").strip())


def _pile_candidates(mark: str) -> list[tuple[str, str]]:
    base = _pile_base(mark)
    names = [
        (f"Сваи {base}", FIELD_PLAIN),
        (base, FIELD_PLAIN),
        (mark, FIELD_PLAIN),
    ]
    if not base.endswith(("у", "У")):
        names.extend(
            (
                (f"Сваи {base}у", FIELD_U),
                (f"{base}у", FIELD_U),
            )
        )
    return [(normalize_name(name), field_name) for name, field_name in names]


def _bridge_candidates(mark: str) -> list[tuple[str, str]]:
    names = [f"Сваи мостовые {mark}", mark]
    matched = _BRIDGE_MARK_RE.match(mark.strip())
    if matched:
        length, section, kind, num = matched.groups()
        if kind in "TТtт":
            names.append(f"Сваи мостовые С {length}.{section}-Т{num}")
        else:
            names.extend(
                (
                    f"Сваи мостовые С {length}.{section}-В{num}",
                    f"Сваи С {length}0.{section}-ВСв.{num}",
                    f"Сваи С {length}0.{section}-НСв.{num}",
                )
            )
    return [(normalize_name(name), FIELD_PLAIN) for name in names]


def _prefixed_candidates(mark: str, prefix: str) -> list[tuple[str, str]]:
    names = [mark]
    stripped = mark.strip()
    if not normalize_name(stripped).startswith(normalize_name(prefix)):
        names.append(f"{prefix} {stripped}")
    return [(normalize_name(name), FIELD_PLAIN) for name in names]


def _step_candidates(mark: str) -> list[tuple[str, str]]:
    names = [mark, f"Лестничные ступени {mark}"]
    matched = _STEP_MARK_RE.match(mark.strip())
    if matched:
        num, rest = matched.groups()
        names.append(f"Лестничные ступени ЛС-{num}{rest.strip()}")
    return [(normalize_name(name), FIELD_PLAIN) for name in names]


def _select_target(
    kind: str,
    norm_name: str,
    targets: Sequence[tuple[str, str]],
) -> tuple[str, str] | None:
    marks = {mark for mark, _field in targets}
    if len(marks) != 1:
        return None
    mark = next(iter(marks))
    fields = [field_name for item_mark, field_name in targets if item_mark == mark]
    unique_fields = list(dict.fromkeys(fields))
    if len(unique_fields) == 1:
        return mark, unique_fields[0]
    if kind == "pile" and FIELD_U in unique_fields and norm_name.endswith("у"):
        return mark, FIELD_U
    return mark, FIELD_PLAIN


def _record_ambiguous(
    kind: str,
    name: str,
    guids: tuple[str, ...],
    targets: Sequence[tuple[str, str]],
    plans: dict[str, _MarkPlan],
    matched_marks: set[str],
    unmatched_amb: list[AmbiguousMatch],
) -> None:
    note = _ambiguous_note(guids)
    marks = {mark for mark, _field in targets}
    if not marks:
        unmatched_amb.append(
            AmbiguousMatch(
                product_kind=kind,
                mark=None,
                name=name,
                guids=guids,
                note=note,
            )
        )
        return
    for mark in sorted(marks):
        matched_marks.add(mark)
        plan = plans[mark]
        plan.ambiguous_names.append(name)
        for guid in guids:
            if guid not in plan.ambiguous_guids:
                plan.ambiguous_guids.append(guid)


def _add_fill(plan: _MarkPlan, field_name: str, guid: str) -> None:
    previous = plan.fills.get(field_name)
    if previous is None:
        plan.fills[field_name] = guid
        return
    if previous != guid:
        plan.fill_conflicts.append(field_name)
        for value in (previous, guid):
            if value not in plan.ambiguous_guids:
                plan.ambiguous_guids.append(value)


def _ambiguous_note(guids: Iterable[str]) -> str:
    listed = ", ".join(guids)
    return f"в 1С несколько GUID: {listed}"


def _persist_ambiguous_candidates(
    conn: sqlite3.Connection,
    kind: str,
    ambiguous: Sequence[AmbiguousMatch],
    incoming: Sequence[PricelistRow],
) -> None:
    if not ambiguous:
        return
    from core.duplicate_candidates import upsert_candidates

    by_guid: dict[str, PricelistRow] = {}
    for row in incoming:
        key = (row.guid or "").strip().lower()
        if key:
            by_guid[key] = row
    for item in ambiguous:
        mark = item.mark or item.name
        rows: list[tuple[str, str, Optional[float]]] = []
        for guid in item.guids:
            src = by_guid.get(str(guid).strip().lower())
            name = src.name if src is not None else item.name
            price = src.price if src is not None else None
            rows.append((guid, name, price))
        if rows:
            upsert_candidates(conn, kind, mark, rows)


def _apply_mark(
    conn: sqlite3.Connection,
    row: NomenclatureGuidRow,
    plan: _MarkPlan,
    *,
    overwrite_manual: bool,
) -> tuple[list[tuple[str, GuidWrite]], Optional[AmbiguousMatch], int]:
    is_manual = row.match_status == "manual" and not overwrite_manual
    is_ambiguous = bool(plan.ambiguous_guids) or bool(plan.fill_conflicts)
    writes: list[tuple[str, GuidWrite]] = []
    unchanged = 0
    fields: dict[str, object] = {}

    for field_name, guid in plan.fills.items():
        if field_name in plan.fill_conflicts:
            continue
        current = getattr(row, field_name)
        current_key = (current or "").strip().lower() or None
        if current_key == guid:
            unchanged += 1
            continue
        if is_manual:
            continue
        if current_key and row.match_status == "manual" and not overwrite_manual:
            continue
        fields[field_name] = guid
        kind_write = "new" if not current_key else "updated"
        writes.append(
            (
                kind_write,
                GuidWrite(
                    product_kind=row.product_kind,
                    mark=row.mark,
                    field=field_name,
                    guid=guid,
                ),
            )
        )

    amb_item: Optional[AmbiguousMatch] = None
    if is_ambiguous:
        note = _ambiguous_note(plan.ambiguous_guids)
        amb_item = AmbiguousMatch(
            product_kind=row.product_kind,
            mark=row.mark,
            name=plan.ambiguous_names[0] if plan.ambiguous_names else row.mark,
            guids=tuple(plan.ambiguous_guids),
            note=note,
        )
        if not is_manual:
            fields["match_status"] = "ambiguous"
            fields["match_note"] = note
    elif writes and not is_manual:
        fields["match_status"] = "auto"
        fields["match_note"] = None

    if fields:
        upsert(conn, row.product_kind, row.mark, **fields)
    return writes, amb_item, unchanged


__all__ = [
    "AmbiguousMatch",
    "DisappearedMark",
    "FIELD_PLAIN",
    "FIELD_U",
    "GuidWrite",
    "SyncReport",
    "Unmatched1C",
    "infer_product_kind",
    "infer_product_kind_from_names",
    "normalize_name",
    "sync_pricelist",
]
