"""Видимое имя карточки договора по справочнику 1С.

Список показывает ``imported_name``, если оно задано, иначе ``full_name``,
иначе ``short_name`` — тот же порядок, что у ``_card_display_name``.
"""

from __future__ import annotations

import sqlite3
from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Iterable, Mapping, Sequence

# 0001/10/26 уже несёт полное имя бланка. Более короткое имя справочника его не затирает.
_BLANK_NUMBER_KEEP_FULLER = "0001/10/26"
_NAME_KEYS = ("imported_name", "full_name", "short_name")


def normalize_directory_name(value: str) -> str:
    """Сравнение имён справочника: регистр, ё/е и лишние пробелы."""
    return " ".join(value.strip().casefold().replace("ё", "е").split())


def choose_directory_name(directory_names: Sequence[str], inn: str) -> str | None:
    """Имя из справочника для этого ИНН, либо None, если выбрать нельзя.

    Одна строка — её имя. Несколько строк с одним именем после нормализации —
    это имя. Если имена разные, строка, чьё имя равно самому ИНН, отбрасывается.
    Одно оставшееся имя берётся. Два разных — None.
    """
    names = _clean_names(directory_names)
    inn_text = inn.strip()
    if not inn_text or not names:
        return None
    if len(names) == 1:
        return names[0]
    grouped = _first_by_normalized(names)
    if len(grouped) == 1:
        return next(iter(grouped.values()))
    without_inn = _first_by_normalized(name for name in names if name != inn_text)
    if len(without_inn) == 1:
        return next(iter(without_inn.values()))
    return None


def visible_card_name(row: Mapping[str, Any]) -> str:
    """Имя, которое показывают список карточек и мастер КП."""
    for key in _NAME_KEYS:
        text = _text(row.get(key))
        if text:
            return text
    return ""


def resolve_card_display_name(
    card_name: str,
    inn: str,
    directory_names: Sequence[str],
    *,
    number: str = "",
    imported_name: str | None = None,
    full_name: str | None = None,
) -> str:
    """Имя, которое должно быть видно. Без однозначного ИНН карточка не меняется."""
    visible = card_name.strip()
    chosen = choose_directory_name(directory_names, inn)
    if chosen is None or chosen == visible:
        return visible
    if _keep_longer_blank_name(
        number=number,
        visible=visible,
        chosen=chosen,
        imported_name=imported_name,
        full_name=full_name,
    ):
        return visible
    return chosen


@dataclass(frozen=True)
class CardNameDecision:
    card_id: int
    number: str
    inn: str
    before: str
    after: str
    reason: str


def plan_card_directory_names(
    cards: Sequence[Mapping[str, Any]],
    names_by_inn: Mapping[str, Sequence[str]],
) -> list[CardNameDecision]:
    """Решение по каждой карточке. ``rename`` — единственная причина писать имя."""
    decisions: list[CardNameDecision] = []
    for card in cards:
        inn = _text(card.get("inn"))
        decisions.append(_decide(card, inn, names_by_inn.get(inn, ())))
    return decisions


def apply_card_directory_names(db_path: str) -> list[CardNameDecision]:
    """Пишет ``imported_name`` одной транзакцией. Справочник и связь не трогает.

    Повтор безопасен: уже совпавшее имя получает причину ``already``.
    """
    conn = sqlite3.connect(db_path, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA busy_timeout=30000")
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("BEGIN IMMEDIATE")
        parties_before = _party_rows(conn)
        links_before = _card_links(conn)
        decisions = plan_card_directory_names(
            _load_cards(conn),
            _names_by_inn(conn),
        )
        updates = [
            (item.after, item.card_id)
            for item in decisions
            if item.reason == "rename"
        ]
        conn.executemany(
            "UPDATE supply_contract SET imported_name = ? WHERE id = ?",
            updates,
        )
        if _party_rows(conn) != parties_before or _card_links(conn) != links_before:
            raise RuntimeError("справочник или связь карточки изменились")
        conn.execute("COMMIT")
        return decisions
    except Exception:
        _rollback_quietly(conn)
        raise
    finally:
        conn.close()


def _decide(
    card: Mapping[str, Any],
    inn: str,
    directory_names: Sequence[str],
) -> CardNameDecision:
    visible = visible_card_name(card)
    number = str(card.get("number") or "")
    imported_name = _optional_text(card.get("imported_name"))
    full_name = _optional_text(card.get("full_name"))
    after = resolve_card_display_name(
        visible,
        inn,
        directory_names,
        number=number,
        imported_name=imported_name,
        full_name=full_name,
    )
    chosen = choose_directory_name(directory_names, inn) if directory_names else None
    reason = _reason(
        directory_names=directory_names,
        chosen=chosen,
        visible=visible,
        after=after,
        number=number,
        imported_name=imported_name,
        full_name=full_name,
    )
    return _decision(card, inn, visible, after, reason)


def _reason(
    *,
    directory_names: Sequence[str],
    chosen: str | None,
    visible: str,
    after: str,
    number: str,
    imported_name: str | None,
    full_name: str | None,
) -> str:
    if not directory_names:
        return "no_directory"
    if chosen is None:
        return "ambiguous"
    if after != visible:
        return "rename"
    if _keep_longer_blank_name(
        number=number,
        visible=visible,
        chosen=chosen,
        imported_name=imported_name,
        full_name=full_name,
    ):
        return "longer_blank"
    return "already"


def _decision(
    card: Mapping[str, Any],
    inn: str,
    before: str,
    after: str,
    reason: str,
) -> CardNameDecision:
    return CardNameDecision(
        card_id=int(card["id"]),
        number=str(card.get("number") or ""),
        inn=inn,
        before=before,
        after=after,
        reason=reason,
    )


def _keep_longer_blank_name(
    *,
    number: str,
    visible: str,
    chosen: str,
    imported_name: str | None,
    full_name: str | None,
) -> bool:
    if number != _BLANK_NUMBER_KEEP_FULLER or _text(imported_name):
        return False
    blank = _text(full_name)
    return bool(blank) and visible == blank and len(blank) > len(chosen)


def _clean_names(directory_names: Sequence[str]) -> list[str]:
    return [name.strip() for name in directory_names if isinstance(name, str) and name.strip()]


def _first_by_normalized(names: Iterable[str]) -> dict[str, str]:
    grouped: dict[str, str] = {}
    for name in names:
        grouped.setdefault(normalize_directory_name(name), name)
    return grouped


def _text(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return value.strip()


def _optional_text(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    return value


def _names_by_inn(conn: sqlite3.Connection) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = defaultdict(list)
    rows = conn.execute(
        """
        SELECT inn, name
        FROM counterparties
        WHERE inn IS NOT NULL AND TRIM(inn) != ''
        ORDER BY id
        """
    )
    for row in rows:
        grouped[_text(row["inn"])].append(str(row["name"]))
    return grouped


def _load_cards(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    rows = conn.execute(
        """
        SELECT id, number, imported_name, full_name, short_name, inn, counterparty_id
        FROM supply_contract
        ORDER BY id
        """
    )
    return [dict(row) for row in rows]


def _party_rows(conn: sqlite3.Connection) -> tuple[tuple[Any, ...], ...]:
    rows = conn.execute(
        """
        SELECT id, code_1c, guid_1c, name, inn, is_client
        FROM counterparties
        ORDER BY id
        """
    )
    return tuple(tuple(row) for row in rows)


def _card_links(conn: sqlite3.Connection) -> tuple[tuple[Any, ...], ...]:
    rows = conn.execute(
        "SELECT id, counterparty_id FROM supply_contract ORDER BY id"
    )
    return tuple(tuple(row) for row in rows)


def _rollback_quietly(conn: sqlite3.Connection) -> None:
    try:
        conn.execute("ROLLBACK")
    except sqlite3.OperationalError:
        pass
