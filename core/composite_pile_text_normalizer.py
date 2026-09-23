#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize composite-pile marks and order text to series canon (С60.30-ВС.1)."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.line_prepare import prepare_source_line

# Prefix «Сваи» / «Свай» before a mark (with optional space after С/C).
_SVAI_PREFIX_RE = re.compile(
    r"^\s*[СC]ваи?\s+",
    re.IGNORECASE | re.UNICODE,
)

# Leading С/C + optional spaces before digits.
_LEADING_C_RE = re.compile(r"^[СC]\s*", re.IGNORECASE | re.UNICODE)

# Section suffix: ВС / НС / ВСв / НСв (+ optional index .N or stuck N).
_SECTION_SUFFIX_RE = re.compile(
    r"-([ВBвbНHнh][СCсc][вv]?\.?)(\d+)?$",
    re.UNICODE,
)

# Whole-pile joint suffixes (longest first when matching).
_JOINT_SUFFIXES = ("-Св.ВП", "-Св", "-С")

# Mark body before grade/qty: starts with С + digits.
_MARK_TOKEN_RE = re.compile(
    r"^([СC]\s*[\d.,]+(?:-[^\s]+)?)",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class CompositePileNormalizeResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)


def _cyrillic_c(ch: str) -> str:
    return "С" if ch.upper() in {"C", "С"} else ch


def _fix_joint_casing(suffix: str) -> str:
    """Normalize joint suffix spelling; keep Cyrillic С."""
    s = suffix.replace("C", "С").replace("c", "С")
    # Latin V/P in ВП → Cyrillic
    s = s.replace("V", "В").replace("v", "В").replace("P", "П").replace("p", "П")
    return s


def _normalize_section_suffix(suffix: str) -> str:
    """`-ВС.1` / `-Всв4` / `-BC.1` → canonical `-ВС.1` / `-ВСв.4`."""
    m = _SECTION_SUFFIX_RE.fullmatch(suffix)
    if not m:
        # Whole-pile joints
        for joint in _JOINT_SUFFIXES:
            if suffix.upper().replace("C", "С") == joint.upper() or _fix_joint_casing(
                suffix
            ) == joint:
                return joint
        # Fallback: Cyrillic С in -С* joints
        fixed = suffix
        if fixed.startswith("-"):
            rest = fixed[1:]
            rest = rest.replace("C", "С").replace("c", "с")
            rest = rest.replace("V", "В").replace("v", "в").replace("P", "П").replace("p", "п")
            # Title-case known joints
            upper = "-" + rest
            for joint in _JOINT_SUFFIXES:
                if upper.upper() == joint.upper():
                    return joint
            return upper
        return suffix

    letters, index = m.group(1), m.group(2)
    # letters like ВС / ВСв / BC / BCv / Всв.
    letters = letters.rstrip(".")
    low = (
        letters.lower()
        .replace("b", "в")
        .replace("h", "н")
        .replace("c", "с")
        .replace("v", "в")
    )
    if low.startswith("всв"):
        canon_letters = "ВСв"
    elif low.startswith("нсв"):
        canon_letters = "НСв"
    elif low.startswith("вс"):
        canon_letters = "ВС"
    elif low.startswith("нс"):
        canon_letters = "НС"
    else:
        canon_letters = letters

    if index is None:
        return f"-{canon_letters}"
    return f"-{canon_letters}.{index}"


def canonicalize_composite_pile_mark(raw: str) -> str:
    """Alias / noisy spelling → series canon mark (no grade/qty).

    Does not touch concrete-grade tokens — pass only the mark fragment.
    """
    text = (raw or "").strip()
    if not text:
        return ""

    text = _SVAI_PREFIX_RE.sub("", text).strip()
    text = re.sub(r"\s+", " ", text)

    # If remainder still has spaces mid-mark («С 60.30-ВС.1»), drop them in mark body.
    # Split off nothing — this function is mark-only.
    text = text.replace(" ", "")

    if not text:
        return ""

    # Leading C/С
    if text[0].upper() in {"C", "С"}:
        text = "С" + text[1:]

    # Split body / suffix at last '-' that starts a known joint or section
    dash = text.rfind("-")
    if dash < 0:
        return text

    body, suffix = text[:dash], text[dash:]
    # Body: latin B in length/section digits area is rare; keep digits/dots.
    # Latin C already fixed at start.
    suffix = _normalize_section_suffix(suffix)
    return body + suffix


def _normalize_order_line(line: str) -> str:
    cleaned = prepare_source_line(line, "composite_piles")
    cleaned = _SVAI_PREFIX_RE.sub("", cleaned).strip()
    match = _MARK_TOKEN_RE.match(cleaned)
    if not match:
        return cleaned
    raw_mark = match.group(1)
    remainder = cleaned[match.end() :].strip()
    canon = canonicalize_composite_pile_mark(raw_mark)
    if remainder:
        return f"{canon} {remainder}"
    return canon


def normalize_composite_pile_order_text(text: str) -> CompositePileNormalizeResult:
    """Split multiline order text; canonicalize marks without rewriting grades."""
    if not text or not text.strip():
        return CompositePileNormalizeResult(normalized_text=text or "")

    raw_lines = [part.strip() for part in re.split(r"[\n;]+", text) if part.strip()]
    normalized_lines = [_normalize_order_line(line) for line in raw_lines]
    return CompositePileNormalizeResult(
        normalized_text="\n".join(normalized_lines),
        normalized_lines=normalized_lines,
    )
