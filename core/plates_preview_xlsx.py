#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""XLSX превью после замены широких плит: ввод → распознано (имя/кол-во) → как в КП (имя/кол-во).

Правила колонок A–E (лист «Превью списка»):

- **A — как прислал пользователь:** одна ячейка на логическую строку ввода из
  ``initial_user_plate_lines``; пустая A — продолжение того же логического блока (раскол ширины,
  например 1,5 м → 1,2 + 0,3) или строка разбора без соответствующей строки ввода (если строк
  ввода меньше, чем логических строк после ``plates_text``).
- **B–C — распознано:** наименование по ключу вклада и построчное количество (вклад этой строки
  ввода в позицию).
- **D–E — как в КП:** то же наименование, что в B, и глобальное количество по ключу по всему
  заказу (снимок ``PLATE_LOAD_DETAILS`` после парсинга ``plates_text``).
- **Схлопывание D–E:** если подряд совпадает тройка ``(текст A после strip, имя D, количество E)``,
  ячейки D–E во второй и далее таких строках остаются пустыми (без дубля подряд).

Раскол одной логической строки на несколько физических позиций выводится отдельными строками
листа; каждая позиция — отдельная строка B–E. Внутри одной логической строки физические строки
упорядочены по возрастанию длины плиты (``length_m``), при равной длине — по ширине, затем по
строке ``length_dm_raw`` из ключа вклада.
"""

from __future__ import annotations

import logging
from pathlib import Path
from collections import OrderedDict
try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    Workbook = None  # type: ignore

import core.config_and_data as cfg
from core.config_and_data import LineContributionKey, make_plate_name, set_plate_lists_from_text
from core.reconciliation_xlsx import split_plate_text_lines

logger = logging.getLogger(__name__)

_HEADER_FONT = Font(bold=True)
_WRAP = Alignment(wrap_text=True, vertical="top")

# A + распознано (имя, кол-во) + КП (имя, кол-во)
_PREVIEW_COLS = 5


def _ldr_equal(a: str, b: str) -> bool:
    return (a or "").strip().replace(",", ".") == (b or "").strip().replace(",", ".")


def qty_for_contribution_key(
    key: LineContributionKey,
    plate_load_details: dict[tuple[float, float, int, str], int],
) -> int:
    """Количество по ключу строки; для раскола 1.5→1.2+0.3 ищет запись с шириной ~1.5 м."""
    lm, wm, lc, ldr = key
    ldr_s = (ldr or "").strip()

    def _load_ok(sload: int) -> bool:
        if lc is None:
            return True
        return cfg.load_code_for_price_match(float(sload)) == cfg.load_code_for_price_match(float(lc))

    for (sl, sw, sload, sldr), q in plate_load_details.items():
        if abs(sl - lm) >= 0.01 or abs(sw - wm) >= 0.01:
            continue
        if not _ldr_equal(ldr_s, sldr):
            continue
        if not _load_ok(sload):
            continue
        return int(q)

    if 1.14 <= wm <= 1.26 or 0.25 <= wm <= 0.34:
        for (sl, sw, sload, sldr), q in plate_load_details.items():
            if abs(sl - lm) >= 0.01:
                continue
            if not (1.45 <= sw <= 1.55):
                continue
            if not _ldr_equal(ldr_s, sldr):
                continue
            if not _load_ok(sload):
                continue
            return int(q)

    return 0


def _sort_contribution_keys(keys: list[LineContributionKey]) -> list[LineContributionKey]:
    """Сортировка по длине (м), ширине (м), затем лексикографически по марке длины в ключе."""

    def _key(k: LineContributionKey) -> tuple[float, float, str]:
        lm, wm, _lc, ldr = k
        return (lm, wm, (ldr or "").strip())

    return sorted(keys, key=_key)


def preview_row_triples_for_contributions(
    keys: list[LineContributionKey],
    line_plate_load_details: dict,
    global_plate_load_details: dict,
    *,
    max_pairs: int | None = None,
) -> list[tuple[str, int, int]]:
    """По уникальным ключам вклада: (наименование, q_line, q_global).

    Уникальные ключи сортируются по длине/ширине (см. ``_sort_contribution_keys``).
    ``max_pairs`` — необязательный лимит числа позиций; ``None`` — все уникальные ключи.
    """
    seen: set[LineContributionKey] = set()
    ordered_unique: list[LineContributionKey] = []
    for key in keys:
        if key in seen:
            continue
        seen.add(key)
        ordered_unique.append(key)

    ordered_unique = _sort_contribution_keys(ordered_unique)

    if max_pairs is not None and len(ordered_unique) > max_pairs:
        logger.warning(
            "[plates_preview] у строки %s уникальных ключей вклада — в вывод только первые %s",
            len(ordered_unique),
            max_pairs,
        )

    to_process = ordered_unique if max_pairs is None else ordered_unique[:max_pairs]

    out: list[tuple[str, int, int]] = []
    for key in to_process:
        lm, wm, lc, ldr = key
        lc_int: int | None
        if lc is not None:
            lc_int = int(round(float(lc)))
        else:
            lc_int = None
        name = make_plate_name(
            lm,
            wm,
            load_code=lc_int,
            length_dm_raw=(ldr or None) if (ldr or "").strip() else None,
        )
        q_line = qty_for_contribution_key(key, line_plate_load_details)
        q_global = qty_for_contribution_key(key, global_plate_load_details)
        out.append((name, q_line, q_global))
    return out


def preview_row_keyed_triples_for_contributions(
    keys: list[LineContributionKey],
    line_plate_load_details: dict,
    global_plate_load_details: dict,
    *,
    max_pairs: int | None = None,
) -> list[tuple[LineContributionKey, str, int, int]]:
    """По уникальным ключам вклада: (key, наименование, q_line, q_global)."""
    seen: set[LineContributionKey] = set()
    ordered_unique: list[LineContributionKey] = []
    for key in keys:
        if key in seen:
            continue
        seen.add(key)
        ordered_unique.append(key)

    ordered_unique = _sort_contribution_keys(ordered_unique)

    if max_pairs is not None and len(ordered_unique) > max_pairs:
        logger.warning(
            "[plates_preview] у строки %s уникальных ключей вклада — в вывод только первые %s",
            len(ordered_unique),
            max_pairs,
        )

    to_process = ordered_unique if max_pairs is None else ordered_unique[:max_pairs]

    out: list[tuple[LineContributionKey, str, int, int]] = []
    for key in to_process:
        lm, wm, lc, ldr = key
        lc_int: int | None
        if lc is not None:
            lc_int = int(round(float(lc)))
        else:
            lc_int = None
        name = make_plate_name(
            lm,
            wm,
            load_code=lc_int,
            length_dm_raw=(ldr or None) if (ldr or "").strip() else None,
        )
        q_line = qty_for_contribution_key(key, line_plate_load_details)
        q_global = qty_for_contribution_key(key, global_plate_load_details)
        out.append((key, name, q_line, q_global))
    return out


def name_qty_pairs_for_contributions(
    keys: list[LineContributionKey],
    plate_load_details: dict[tuple[float, float, int, str], int],
    *,
    max_pairs: int | None = None,
) -> list[tuple[str, int]]:
    """Пары (наименование, кол-во) по уникальным ключам вклада (один словарь для количества)."""
    triples = preview_row_triples_for_contributions(
        keys, plate_load_details, plate_load_details, max_pairs=max_pairs
    )
    return [(a, b) for a, b, _ in triples]


def build_plates_reconciliation_preview_xlsx(
    path: str | Path,
    *,
    plates_text: str,
    initial_user_plate_lines: list[str],
) -> None:
    """
    Парсит ``plates_text``, строит лист превью с колонками A–E (см. модульный докстринг).

    Для каждой логической строки выводятся все уникальные ключи вклада (физические строки B–E);
    колонка A заполняется только в первой строке блока, далее — пусто до следующей логической строки.
    """
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl не установлен — превью XLSX недоступно")

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    _, line_contributions, line_plate_load_details = set_plate_lists_from_text(plates_text)
    details_global = dict(cfg.PLATE_LOAD_DETAILS)

    norm_lines = split_plate_text_lines(plates_text)
    n = max(len(norm_lines), len(line_contributions), len(line_plate_load_details))
    if len(norm_lines) != len(line_contributions) or len(norm_lines) != len(line_plate_load_details):
        logger.warning(
            "[plates_preview] len(norm_lines)=%s line_contributions=%s line_plate_load_details=%s",
            len(norm_lines),
            len(line_contributions),
            len(line_plate_load_details),
        )

    nu = len(initial_user_plate_lines)
    if nu != n:
        logger.warning(
            "[plates_preview] len(initial_user_plate_lines)=%s != строк превью=%s",
            nu,
            n,
        )

    # Сырые строки вклада (по строкам ввода)
    contribution_rows: list[dict[str, object]] = []
    no_match_rows: list[tuple[int, str]] = []

    # Физические строки: (col_a, name_b, qty_c, kp_name, kp_qty_int)
    physical: list[tuple[str, str, int | None, str, int]] = []
    for i in range(n):
        user_cell = (
            initial_user_plate_lines[i]
            if i < len(initial_user_plate_lines)
            else ""
        )
        contrib = line_contributions[i] if i < len(line_contributions) else []
        line_det = line_plate_load_details[i] if i < len(line_plate_load_details) else {}
        keyed_triples = preview_row_keyed_triples_for_contributions(
            contrib, line_det, details_global
        )
        if not keyed_triples:
            no_match_rows.append((i, user_cell))
            continue

        for key, name, q_line, q_global in keyed_triples:
            contribution_rows.append(
                {
                    "line_idx": i,
                    "user_cell": user_cell,
                    "key": key,
                    "name": name,
                    "q_line": q_line,
                    "q_global": q_global,
                }
            )

    # Группируем вклад по позиции КП: D–E показываем один раз в начале группы.
    grouped: OrderedDict[tuple[LineContributionKey, str, int], list[dict[str, object]]] = OrderedDict()
    for row in contribution_rows:
        kp_name = str(row["name"])
        kp_qty = int(row["q_global"])
        key = row["key"]
        group_key = (key, kp_name, kp_qty)
        grouped.setdefault(group_key, []).append(row)

    for (_k, kp_name, kp_qty), rows in grouped.items():
        for idx, row in enumerate(rows):
            a = str(row["user_cell"])
            b = str(row["name"])
            q_line = int(row["q_line"])
            c_val = None if q_line <= 0 else q_line
            if idx == 0:
                physical.append((a, b, c_val, kp_name, kp_qty))
            else:
                physical.append((a, b, c_val, "", 0))

    # Нераспознанные/пустые строки без вкладов сохраняем в хвосте, чтобы не терять связь с вводом.
    for _line_idx, user_cell in no_match_rows:
        physical.append((user_cell, "", "", "", 0))

    wb = Workbook()
    ws = wb.active
    ws.title = "Превью списка"

    headers = [
        "Как прислал пользователь",
        "Распознано (наименование)",
        "Распознано (кол-во)",
        "Как в КП (наименование)",
        "КП (кол-во)",
    ]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = _HEADER_FONT
        cell.alignment = _WRAP

    for idx, (a, b, c, kp_name, kp_qty) in enumerate(physical):
        r = idx + 2
        ws.cell(row=r, column=1, value=a).alignment = _WRAP
        ws.cell(row=r, column=2, value=b).alignment = _WRAP
        ws.cell(row=r, column=3, value=c).alignment = _WRAP

        d_val = kp_name if kp_name else ""
        e_val = None if kp_qty <= 0 else kp_qty

        ws.cell(row=r, column=4, value=d_val).alignment = _WRAP
        ws.cell(row=r, column=5, value=e_val).alignment = _WRAP

    for col in range(1, _PREVIEW_COLS + 1):
        ws.column_dimensions[get_column_letter(col)].width = 28 if col == 1 else 22

    wb.save(str(path))
