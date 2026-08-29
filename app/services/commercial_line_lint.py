"""Per-physical-line lint of commercial source text.

Uses line parsers only. Does not run preview, ILP, or draft persistence.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from core.bridge_pile_line_parser import parse_bridge_pile_line
from core.fbs_line_parser import parse_fbs_line
from core.march_line_parser import parse_march_line
from core.pile_line_parser import parse_pile_line
from core.plate_line_parser import parse_line
from core.plate_validation import validate_plate_values
from core.step_line_parser import parse_step_line


@dataclass(frozen=True)
class LineLint:
    index: int
    text: str
    empty: bool
    ok: bool
    reason_text: str | None = None


def _lint_plate_line(raw: str) -> tuple[bool, str | None]:
    parsed = parse_line(raw)
    if not parsed.parsed:
        return False, parsed.reason_text or "строка не распознана"
    validation = validate_plate_values(parsed.width_m, parsed.length_m, parsed.qty)
    if not validation.ok:
        return False, validation.reason_text or "значения строки некорректны"
    return True, None


def _lint_generic(parse_fn: Callable[..., object]) -> Callable[[str], tuple[bool, str | None]]:
    def _check(raw: str) -> tuple[bool, str | None]:
        result = parse_fn(raw)
        parsed = bool(getattr(result, "parsed", False))
        if parsed:
            return True, None
        reason = getattr(result, "reason_text", None)
        return False, str(reason) if reason else "строка не распознана"

    return _check


_GENERIC_LINTERS: dict[str, Callable[[str], tuple[bool, str | None]]] = {
    "piles": _lint_generic(parse_pile_line),
    "steps": _lint_generic(parse_step_line),
    "marches": _lint_generic(parse_march_line),
    "bridge_piles": _lint_generic(parse_bridge_pile_line),
    "fbs": _lint_generic(parse_fbs_line),
}


def lint_source_lines(text: str, product_type: str) -> list[LineLint]:
    """Lint each physical ``\\n`` line. Blank / whitespace-only lines are ok."""
    physical_lines = text.split("\n")
    results: list[LineLint] = []
    for index, raw in enumerate(physical_lines):
        if not raw.strip():
            results.append(LineLint(index=index, text=raw, empty=True, ok=True, reason_text=None))
            continue
        if product_type == "plates":
            ok, reason = _lint_plate_line(raw)
        else:
            linter = _GENERIC_LINTERS.get(product_type)
            if linter is None:
                raise ValueError(f"Unsupported product_type: {product_type}")
            ok, reason = linter(raw)
        results.append(LineLint(index=index, text=raw, empty=False, ok=ok, reason_text=None if ok else reason))
    return results


def unparsed_line_texts(lines: list[LineLint]) -> list[str]:
    return [line.text for line in lines if not line.empty and not line.ok]
