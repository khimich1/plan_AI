from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from app.domain.models.plate_order import PlateOrder


@dataclass
class ParseResult:
    order: PlateOrder
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)
    unparsed_lines: list[str] = field(default_factory=list)
    diagnostics: list[dict[str, Any]] = field(default_factory=list)
    wide_plate_lines: list[tuple[str, int]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    line_contributions: list[list[tuple[float, float, float | None, str]]] = field(default_factory=list)
    line_plate_load_details: list[dict[tuple, int]] = field(default_factory=list)

