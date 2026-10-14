from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from core.commercial_pricing import lookup_composite_pile_price
from core.composite_pile_line_parser import expand_order_line, merge_composite_pile_sections
from core.composite_pile_text_normalizer import normalize_composite_pile_order_text
from core.exceptions import PriceNotFoundError


@dataclass
class CommercialCompositePilePreviewResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)
    unparsed_lines: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    order_data: list[dict[str, Any]] = field(default_factory=list)
    total_sum: float = 0.0


class CommercialCompositePileService:
    """Build composite-pile order_data: expand → canon sections → price lookup."""

    def generate_preview(
        self,
        text: str,
        *,
        db_path: str,
        default_grade: str = "B25",
    ) -> CommercialCompositePilePreviewResult:
        normalized = normalize_composite_pile_order_text(text)
        raw_lines = normalized.normalized_lines or [
            part.strip() for part in re.split(r"[\n;]+", text or "") if part.strip()
        ]

        sections = []
        unparsed_lines: list[str] = []
        for line in raw_lines:
            result = expand_order_line(line, db_path=db_path, default_grade=default_grade)
            if not result.parsed:
                unparsed_lines.append(line)
                continue
            sections.extend(result.sections)

        merged = merge_composite_pile_sections(sections, default_grade=default_grade)

        order_data: list[dict[str, Any]] = []
        total_sum = 0.0
        for item in merged:
            unit_price: float | None
            try:
                unit_price = lookup_composite_pile_price(
                    item.mark,
                    item.concrete_grade or default_grade,
                    db_path=db_path,
                )
            except PriceNotFoundError:
                unit_price = None

            qty = int(item.qty or 0)
            line_total = (unit_price or 0.0) * qty
            total_sum += line_total
            order_data.append(
                {
                    "product_kind": "composite_pile",
                    "product_type": "composite_piles",
                    "name": item.mark,
                    "mark": item.mark,
                    "concrete_grade": item.concrete_grade or default_grade,
                    "qty": qty,
                    "unit_price": unit_price,
                    "line_total": line_total if unit_price is not None else None,
                }
            )

        return CommercialCompositePilePreviewResult(
            normalized_text=normalized.normalized_text,
            normalized_lines=normalized.normalized_lines,
            unparsed_lines=unparsed_lines,
            order_data=order_data,
            total_sum=total_sum,
        )
