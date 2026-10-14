#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse / expand composite-pile order lines into canonical section rows."""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from typing import Optional

from core.composite_pile_kit import get_composite_pile_assembly, list_section_marks
from core.composite_pile_text_normalizer import (
    canonicalize_composite_pile_mark,
    normalize_composite_pile_order_text,
)
from core.pile_price_db import GRADE_CODES, grade_code_from_value
from core.price_db import DEFAULT_DB

DEFAULT_CONCRETE_GRADE = "B25"

_SECTION_MARK_RE = re.compile(
    r"^С\d+[.,]\d+-(?:ВС|НС|ВСв|НСв)(?:\.\d+)?$",
    re.UNICODE,
)

_MARK_HEAD_RE = re.compile(
    r"^([СC]\s*[\d.,]+(?:-[^\s]+)?)",
    re.IGNORECASE | re.UNICODE,
)


@dataclass(frozen=True)
class CompositePileSection:
    mark: str
    concrete_grade: str
    qty: int
    parsed: bool = True
    reason_code: str = ""
    reason_text: str = ""


@dataclass
class CompositeExpandResult:
    """One order line → zero, one, or two canonical sections."""

    parsed: bool
    sections: list[CompositePileSection] = field(default_factory=list)
    raw_mark: str = ""
    reason_code: str = ""
    reason_text: str = ""


def _parse_grade_and_qty(
    tokens: list[str],
    *,
    default_grade: str,
) -> tuple[str, int] | None:
    if not tokens:
        return default_grade, 1

    if len(tokens) == 1:
        token = tokens[0]
        grade = grade_code_from_value(token)
        if grade:
            return grade, 1
        if token.isdigit():
            return default_grade, max(1, int(token))
        return None

    last = tokens[-1]
    if not last.isdigit():
        return None

    qty = max(1, int(last))
    grade_tokens = tokens[:-1]
    grade_text = " ".join(grade_tokens)
    grade = grade_code_from_value(grade_text) or grade_code_from_value(grade_tokens[0])
    if grade is None:
        return None
    return grade, qty


def is_composite_section_mark(canon: str) -> bool:
    """True if mark is a series section (ВС/НС/ВСв/НСв), not a whole pile."""
    return bool(_SECTION_MARK_RE.match(canon or ""))


# Back-compat alias for internal callers.
_is_section_mark = is_composite_section_mark


def expand_order_line(
    raw: str,
    *,
    db_path: str = DEFAULT_DB,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> CompositeExpandResult:
    """Одна строка заявки → одна секция или две секции комплектации.

    Марки в результате уже канон. Целая марка вне таблицы комплектации
    не раскладывается.
    """
    line = (raw or "").strip()
    if not line:
        return CompositeExpandResult(
            parsed=False,
            reason_code="empty_line",
            reason_text="пустая строка",
        )

    normalized = normalize_composite_pile_order_text(line)
    prepared = (normalized.normalized_lines or [line])[0]
    match = _MARK_HEAD_RE.match(prepared)
    if not match:
        return CompositeExpandResult(
            parsed=False,
            reason_code="pattern_not_matched",
            reason_text="не совпал формат строки составной сваи",
        )

    raw_mark = match.group(1)
    canon = canonicalize_composite_pile_mark(raw_mark)
    remainder = prepared[match.end() :].strip()
    tokens = remainder.split() if remainder else []
    parsed_gq = _parse_grade_and_qty(tokens, default_grade=default_grade)
    if parsed_gq is None:
        return CompositeExpandResult(
            parsed=False,
            raw_mark=canon,
            reason_code="grade_qty_parse_failed",
            reason_text="не удалось распознать класс бетона или количество",
        )
    grade, qty = parsed_gq
    if grade not in GRADE_CODES:
        return CompositeExpandResult(
            parsed=False,
            raw_mark=canon,
            reason_code="unknown_grade",
            reason_text=f"неизвестный класс бетона: {grade}",
        )

    assembly = get_composite_pile_assembly(canon, db_path)
    if assembly is not None:
        upper, lower = assembly
        return CompositeExpandResult(
            parsed=True,
            raw_mark=canon,
            sections=[
                CompositePileSection(mark=upper, concrete_grade=grade, qty=qty),
                CompositePileSection(mark=lower, concrete_grade=grade, qty=qty),
            ],
        )

    if _is_section_mark(canon):
        return CompositeExpandResult(
            parsed=True,
            raw_mark=canon,
            sections=[
                CompositePileSection(mark=canon, concrete_grade=grade, qty=qty),
            ],
        )

    # Known section from kit but regex miss — still accept if in kit sections
    if canon in list_section_marks(db_path):
        return CompositeExpandResult(
            parsed=True,
            raw_mark=canon,
            sections=[
                CompositePileSection(mark=canon, concrete_grade=grade, qty=qty),
            ],
        )

    return CompositeExpandResult(
        parsed=False,
        raw_mark=canon,
        reason_code="pattern_not_matched",
        reason_text="марка не найдена в комплектации составных свай",
    )


def merge_composite_pile_sections(
    sections: list[CompositePileSection],
    *,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[CompositePileSection]:
    """Merge sections with the same canon mark + concrete grade (sum qty)."""
    merged: dict[tuple[str, str], CompositePileSection] = {}
    for section in sections:
        if not section.parsed:
            continue
        grade = section.concrete_grade or default_grade
        key = (section.mark, grade)
        if key in merged:
            existing = merged[key]
            merged[key] = replace(existing, qty=existing.qty + section.qty)
        else:
            merged[key] = replace(section, concrete_grade=grade)
    return list(merged.values())


def expand_composite_pile_order_text(
    text: str,
    *,
    db_path: str = DEFAULT_DB,
    default_grade: str = DEFAULT_CONCRETE_GRADE,
) -> list[CompositePileSection]:
    """Normalize + expand multiline text; merge identical canon+grade rows.

    Unparsed lines are dropped from the merged list (caller should track them
    via expand_order_line per raw line if needed).
    """
    normalized = normalize_composite_pile_order_text(text)
    lines = normalized.normalized_lines or [
        part.strip() for part in re.split(r"[\n;]+", text or "") if part.strip()
    ]
    sections: list[CompositePileSection] = []
    for line in lines:
        result = expand_order_line(line, db_path=db_path, default_grade=default_grade)
        if result.parsed:
            sections.extend(result.sections)
        else:
            sections.append(
                CompositePileSection(
                    mark=result.raw_mark or "",
                    concrete_grade=default_grade,
                    qty=0,
                    parsed=False,
                    reason_code=result.reason_code,
                    reason_text=result.reason_text,
                )
            )
    parsed_only = [s for s in sections if s.parsed]
    unparsed = [s for s in sections if not s.parsed]
    return merge_composite_pile_sections(parsed_only, default_grade=default_grade) + unparsed
