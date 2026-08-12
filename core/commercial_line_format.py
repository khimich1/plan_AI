"""Format commercial offer line display names (mark + optional concrete grade)."""

from __future__ import annotations

from typing import Any, Mapping


def format_line_name(item: Mapping[str, Any]) -> str:
    """Build line name: mark (or name) with optional `` (grade)`` suffix.

    Grade is appended only when ``concrete_grade`` is present and non-empty
    after strip. Empty/missing mark returns a safe string (no crash).
    """
    base = str(item.get("mark") or item.get("name") or "").strip()
    grade = str(item.get("concrete_grade") or "").strip()
    if not grade:
        return base
    if not base:
        return f"({grade})"
    return f"{base} ({grade})"
