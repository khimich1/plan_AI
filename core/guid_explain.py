#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""«Почему нет GUID»: причина, та же карточка, срез своего вида, поиск.

Только чтение pb.db. Ничего не подставляет и не пишет: человек выбирает сам.
Спека: ai_docs/specs/pochemu-net-guid.md (D-same-card, D-slice, D-label, D-search).
"""

from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from typing import Callable, Optional

from core.bridge_pile_price_db import normalize_bridge_pile_mark_for_lookup
from core.fbs_price_db import normalize_fbs_mark_for_lookup
from core.guid_gate import (
    INVOICE_STATUSES,
    OrderLine,
    _normalize_kind,
    check_invoice_guids,
)
from core.kp_db_nomenclature import PLATE_DUP, PLATE_FOUND, find_plate_card
from core.march_price_db import normalize_march_mark
from core.nomenclature_guid import NomenclatureGuidRow
from core.pile_catalog import (
    normalize_pile_mark_key,
    parse_bridge_pile_geometry,
    parse_pile_mark,
    strip_pile_load_suffix,
)
from core.pile_price_db import strip_trailing_u_suffix
from core.step_price_db import normalize_step_mark

SLICE_LIMIT = 30
SEARCH_MIN_CHARS = 3
ROLE_MANAGER = "manager"
ROLE_ECONOMIST = "economist"

LABEL_ZERO = "лишний ноль"
LABEL_LETTER = "другая буква"
LABEL_WIDTH = "другая ширина"
LABEL_LOAD = "другая нагрузка"

_PRICE_TABLES = {
    "pile": "pile_prices",
    "bridge_pile": "bridge_pile_prices",
    "composite_pile": "composite_pile_prices",
    "fbs": "fbs_prices",
    "stair_flight": "march_prices",
    "stair_step": "step_prices",
}
_PLATE_RE = re.compile(
    r"П[БК]\s*(?P<len>[\d,.]+)-(?P<wid>[\d,.]+)-(?P<load>[\d,.]+)п(?P<tail>.*)$",
    re.IGNORECASE,
)
_FBS_RE = re.compile(r"ФБС\s*(\d+)\.(\d+)\.(\d+)", re.IGNORECASE)
_STEP_SERIES_RE = re.compile(r"^(ЛС\d+)")


@dataclass(frozen=True)
class GuidCard:
    """Одна карточка 1С: имя, как в выгрузке, и её GUID."""

    name: str
    guid: str


@dataclass(frozen=True)
class GuidCandidate:
    """Строка среза/поиска: марка, сколько номеров 1С у неё есть, подпись различия.

    cards заполнен только когда номеров два и больше: кнопки берут имя и GUID отсюда.
    Один номер сюда не кладём — его подставляет сервер при замене.
    """

    mark: str
    guid_count: int
    label: Optional[str] = None
    cards: tuple[GuidCard, ...] = ()

    def selectable_for(self, role: str) -> bool:
        """Менеджер нажимает только строку с одним номером; экономист — любую."""
        if role == ROLE_MANAGER:
            return self.guid_count == 1
        return role == ROLE_ECONOMIST


@dataclass(frozen=True)
class GuidExplain:
    product_kind: str
    mark: str
    reason: str  # "" — номер у строки есть
    hint: str
    same_cards: tuple[GuidCandidate, ...] = ()
    slice: tuple[GuidCandidate, ...] = ()

    @property
    def same_card(self) -> Optional[GuidCandidate]:
        return self.same_cards[0] if self.same_cards else None


# ------------------------------------------------------------------ общее


def _num(text: str) -> Optional[float]:
    try:
        return float(str(text).replace(",", "."))
    except ValueError:
        return None


def _fold_text(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "").casefold())


def _letter_label(typed: str, other: str, fold: Callable[[str], str]) -> Optional[str]:
    """«Другая буква»: марки равны после свёртки C↔С, T↔Т, регистра, пробелов."""
    if typed != other and fold(typed) == fold(other):
        return LABEL_LETTER
    return None


# ------------------------------------------------------------------ плиты


@dataclass(frozen=True)
class _PlateParts:
    length: float
    width: float
    load: float
    raw: tuple[str, str, str]
    tail: str


def _parse_plate(name: str) -> Optional[_PlateParts]:
    match = _PLATE_RE.search(str(name or ""))
    if not match:
        return None
    length, width, load = (_num(match.group(k)) for k in ("len", "wid", "load"))
    if length is None or width is None or load is None:
        return None
    raw = (match.group("len"), match.group("wid"), match.group("load"))
    return _PlateParts(length, width, load, raw, match.group("tail").strip())


def _plate_label(typed: str, other: str) -> Optional[str]:
    """Подпись только при одном отличии: хвостовой ноль, одна ширина или одна нагрузка."""
    a, b = _parse_plate(typed), _parse_plate(other)
    if a is None or b is None or a.tail != b.tail:
        return None
    diffs: list[str] = []
    values = ((a.length, b.length, None), (a.width, b.width, LABEL_WIDTH), (a.load, b.load, LABEL_LOAD))
    for index, (left, right, numeric_label) in enumerate(values):
        if a.raw[index] == b.raw[index]:
            continue
        if abs(left - right) < 1e-9:
            diffs.append(LABEL_ZERO)
        elif numeric_label is None:
            return None  # другая длина — не из списка подписей
        else:
            diffs.append(numeric_label)
    return diffs[0] if len(diffs) == 1 else None


def _choice(cards: tuple[GuidCard, ...]) -> tuple[GuidCard, ...]:
    """Кнопки только при двух и более карточках. Один GUID клиенту не отдаём."""
    return cards if len(cards) >= 2 else ()


def _plate_index(conn: sqlite3.Connection) -> dict[str, tuple[GuidCard, ...]]:
    """Имя карточки → её GUID в порядке идентификатора. Пустой GUID не считается."""
    grouped: dict[str, set[str]] = {}
    cur = conn.execute(
        'SELECT "Товар", "Уникальный идентификатор (Номенклатура)" FROM prays_plity'
    )
    for name, guid in cur:
        if name:
            grouped.setdefault(str(name), set()).add(str(guid or "").strip())
    index: dict[str, tuple[GuidCard, ...]] = {}
    for name, guids in grouped.items():
        index[name] = tuple(
            GuidCard(name=name, guid=guid) for guid in sorted(item for item in guids if item)
        )
    return index


def _plate_same_card(
    conn: sqlite3.Connection, typed: str, index: dict[str, tuple[GuidCard, ...]]
) -> Optional[GuidCandidate]:
    match = find_plate_card(conn.cursor(), typed)
    if match.status not in (PLATE_FOUND, PLATE_DUP) or not match.canonical_name:
        return None
    if match.canonical_name == typed:
        return None
    cards = index.get(match.canonical_name, ())
    return GuidCandidate(
        mark=match.canonical_name,
        guid_count=match.guid_count,
        label=_plate_label(typed, match.canonical_name),
        cards=_choice(cards) if match.guid_count >= 2 else (),
    )


def _plate_candidate(typed: str, name: str, cards: tuple[GuidCard, ...]) -> GuidCandidate:
    return GuidCandidate(
        mark=name,
        guid_count=len(cards),
        label=_plate_label(typed, name),
        cards=_choice(cards),
    )


def _plate_slice(
    typed: str, index: dict[str, tuple[GuidCard, ...]], skip: set[str]
) -> tuple[GuidCandidate, ...]:
    parts = _parse_plate(typed)
    if parts is None:
        return ()
    scored: list[tuple[float, float, str, tuple[GuidCard, ...]]] = []
    for name, cards in index.items():
        other = _parse_plate(name)
        if other is None or name in skip or abs(other.length - parts.length) > 0.05:
            continue
        scored.append((abs(other.width - parts.width), abs(other.load - parts.load), name, cards))
    scored.sort(key=lambda item: (item[0], item[1], item[2]))
    return tuple(
        _plate_candidate(typed, name, cards) for _w, _l, name, cards in scored[:SLICE_LIMIT]
    )


# ------------------------------------------------------------------ не плиты


def _fold_for(kind: str) -> Callable[[str], str]:
    if kind == "pile" or kind == "composite_pile":
        return lambda mark: normalize_pile_mark_key(strip_trailing_u_suffix(mark))
    if kind == "bridge_pile":
        return normalize_bridge_pile_mark_for_lookup
    if kind == "fbs":
        return normalize_fbs_mark_for_lookup
    if kind == "stair_step":
        return normalize_step_mark
    return lambda mark: _fold_text(normalize_march_mark(mark)).upper()


def _slice_key_for(kind: str) -> Callable[[str], Optional[object]]:
    """Ось среза: сечение сваи, ширина и высота ФБС, серия ступени/марша."""
    if kind in ("pile", "composite_pile"):
        return lambda mark: parse_pile_mark(mark)[1]
    if kind == "bridge_pile":
        return lambda mark: parse_bridge_pile_geometry(mark)[1]
    if kind == "fbs":
        return _fbs_key
    if kind == "stair_step":
        return _step_key
    return _march_key


def _fbs_key(mark: str) -> Optional[tuple[str, str]]:
    match = _FBS_RE.search(str(mark or ""))
    return (match.group(2), match.group(3)) if match else None


def _step_key(mark: str) -> Optional[str]:
    match = _STEP_SERIES_RE.match(normalize_step_mark(mark))
    return match.group(1) if match else None


def _march_key(mark: str) -> Optional[str]:
    series = normalize_march_mark(mark).split("-")[0]
    return _fold_text(series).upper() if "-" in normalize_march_mark(mark) else None


def _load_universe(
    conn: sqlite3.Connection, kind: str
) -> dict[str, Optional[NomenclatureGuidRow]]:
    """Марки, которые программа знает: справочник номеров и прайс вида."""
    universe: dict[str, Optional[NomenclatureGuidRow]] = {}
    for mark in _price_marks(conn, kind):
        universe[mark] = None
    cur = conn.execute(
        "SELECT product_kind, mark, guid_1c, guid_1c_u, match_status, match_note, updated_at "
        "FROM nomenclature_guid WHERE product_kind = ?",
        (kind,),
    )
    for values in cur:
        universe[values[1]] = NomenclatureGuidRow(*values)
    return universe


def _price_marks(conn: sqlite3.Connection, kind: str) -> list[str]:
    table = _PRICE_TABLES[kind]
    try:
        rows = conn.execute(f"SELECT DISTINCT mark FROM {table}").fetchall()
    except sqlite3.OperationalError:
        return []
    return [str(row[0]).strip() for row in rows if row[0] and str(row[0]).strip()]


def _guid_count(row: Optional[NomenclatureGuidRow], *, with_u: bool) -> int:
    if row is None:
        return 0
    if row.match_status == "ambiguous":
        return 2
    if row.match_status not in INVOICE_STATUSES:
        return 0
    guid = row.guid_1c_u if with_u else row.guid_1c
    return 1 if str(guid or "").strip() else 0


def _non_plate_label(kind: str, typed: str, other: str, fold) -> Optional[str]:
    if kind == "pile":
        typed_u = typed != strip_trailing_u_suffix(typed)
        other_u = other != strip_trailing_u_suffix(other)
        if typed_u != other_u:
            return None
        stems = [normalize_pile_mark_key(strip_pile_load_suffix(strip_trailing_u_suffix(m))) for m in (typed, other)]
        if stems[0] == stems[1] and fold(typed) != fold(other):
            return LABEL_LOAD
    return _letter_label(typed, other, fold)


def _spellings(kind: str, mark: str, row: Optional[NomenclatureGuidRow], *, always_u: bool) -> list[GuidCandidate]:
    """Строки марки: у сваи — без «у» и с «у» (номер «у» — только guid_1c_u)."""
    if kind != "pile":
        return [GuidCandidate(mark=mark, guid_count=_guid_count(row, with_u=False))]
    base = strip_trailing_u_suffix(mark)
    out = [GuidCandidate(mark=base, guid_count=_guid_count(row, with_u=False))]
    if always_u or _guid_count(row, with_u=True):
        out.append(GuidCandidate(mark=f"{base}у", guid_count=_guid_count(row, with_u=True)))
    return out


def _with_label(item: GuidCandidate, kind: str, typed: str, fold) -> GuidCandidate:
    return GuidCandidate(
        mark=item.mark,
        guid_count=item.guid_count,
        label=_non_plate_label(kind, typed, item.mark, fold),
        cards=item.cards,
    )


def _duplicate_cards(conn: sqlite3.Connection, kind: str, mark: str) -> tuple[GuidCard, ...]:
    """Имена карточек 1С из duplicate_candidates. Таблицу не создаём: объяснение только читает."""
    exists = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'duplicate_candidates'"
    ).fetchone()
    if exists is None:
        return ()
    keys = [mark]
    if kind == "pile":
        base = strip_trailing_u_suffix(mark)
        if base and base != mark:
            keys.append(base)
    found: list[GuidCard] = []
    seen: set[str] = set()
    for key in keys:
        rows = conn.execute(
            "SELECT guid, name FROM duplicate_candidates "
            "WHERE active = 1 AND kind = ? AND mark_norm = ? ORDER BY guid",
            (kind, key),
        ).fetchall()
        for guid, name in rows:
            text = str(guid or "").strip()
            if not text or text in seen:
                continue
            seen.add(text)
            found.append(GuidCard(name=str(name or "").strip() or key, guid=text))
        if len(found) >= 2:
            break
    return _choice(tuple(found))


def _with_cards(conn: sqlite3.Connection, kind: str, item: GuidCandidate) -> GuidCandidate:
    if item.guid_count < 2 or item.cards:
        return item
    cards = _duplicate_cards(conn, kind, item.mark)
    if not cards:
        return item
    return GuidCandidate(
        mark=item.mark, guid_count=item.guid_count, label=item.label, cards=cards
    )


def _explain_non_plate(conn: sqlite3.Connection, kind: str, typed: str):
    fold = _fold_for(kind)
    slice_key = _slice_key_for(kind)
    universe = _load_universe(conn, kind)
    typed_fold = fold(typed)
    typed_key = slice_key(typed)
    same: list[GuidCandidate] = []
    others: list[GuidCandidate] = []
    for mark, row in sorted(universe.items()):
        if fold(mark) == typed_fold:
            same.extend(_spellings(kind, mark, row, always_u=True))
        elif typed_key is not None and slice_key(mark) == typed_key:
            others.extend(_spellings(kind, mark, row, always_u=False))
    same = [
        _with_cards(conn, kind, _with_label(item, kind, typed, fold))
        for item in same
        if kind == "pile" or item.mark != typed
    ]
    others = [
        _with_cards(conn, kind, _with_label(item, kind, typed, fold)) for item in others
    ]
    return tuple(same), tuple(others[:SLICE_LIMIT])


# ------------------------------------------------------------------ вход


def explain_missing_guid(line: OrderLine, conn: sqlite3.Connection) -> GuidExplain:
    """Та же карточка по правилу вида, срез своего вида, число номеров.

    Подпись различия — только если отличие одно: буква, хвостовой ноль, ширина или
    нагрузка. Иначе подписи нет, строки отдаются как есть.
    """
    report = check_invoice_guids([line], conn)
    reason, hint = "", ""
    if report.missing:
        reason, hint = report.missing[0].reason, report.missing[0].action_hint
    kind = _canonical_kind(line)
    typed = line.mark
    if kind == "plate":
        index = _plate_index(conn)
        same = _plate_same_card(conn, typed, index)
        skip = {typed} | ({same.mark} if same else set())
        return GuidExplain(
            kind, typed, reason, hint,
            same_cards=(same,) if same else (),
            slice=_plate_slice(typed, index, skip),
        )
    same_cards, slice_rows = _explain_non_plate(conn, kind, typed)
    skip = {item.mark for item in same_cards} | {typed}
    slice_rows = tuple(item for item in slice_rows if item.mark not in skip)
    return GuidExplain(kind, typed, reason, hint, same_cards=same_cards, slice=slice_rows)


def _canonical_kind(line: OrderLine) -> str:
    return _normalize_kind(line.product_kind)


def search_guid_marks(
    conn: sqlite3.Connection, product_kind: str, query: str
) -> tuple[GuidCandidate, ...]:
    """Поиск по марке своего вида: от трёх знаков, до 30 строк, номера считаем честно."""
    text = str(query or "").strip()
    if len(text) < SEARCH_MIN_CHARS:
        return ()
    kind = _canonical_kind(OrderLine(product_kind, text))
    if kind == "plate":
        needle = _fold_text(text)
        found = [
            GuidCandidate(mark=name, guid_count=len(cards), cards=_choice(cards))
            for name, cards in _plate_index(conn).items()
            if needle in _fold_text(name)
        ]
    else:
        fold = _fold_for(kind)
        needles = {_fold_text(text), fold(text)}
        found = []
        for mark, row in _load_universe(conn, kind).items():
            if any(n and (n in _fold_text(mark) or n in fold(mark)) for n in needles):
                found.extend(
                    _with_cards(conn, kind, item)
                    for item in _spellings(kind, mark, row, always_u=False)
                )
    found.sort(key=lambda item: (not _fold_text(item.mark).startswith(_fold_text(text)), item.mark))
    return tuple(found[:SLICE_LIMIT])


__all__ = [
    "GuidCard",
    "GuidCandidate",
    "GuidExplain",
    "ROLE_ECONOMIST",
    "ROLE_MANAGER",
    "SEARCH_MIN_CHARS",
    "SLICE_LIMIT",
    "explain_missing_guid",
    "search_guid_marks",
]
