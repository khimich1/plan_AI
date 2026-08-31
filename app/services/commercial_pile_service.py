from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from core.commercial_pricing import lookup_pile_price
from core.exceptions import PriceNotFoundError
from core.pile_line_parser import parse_pile_line, merge_pile_lines
from core.pile_text_normalizer import normalize_pile_order_text


@dataclass
class CommercialPilePreviewResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)
    unparsed_lines: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    order_data: list[dict[str, Any]] = field(default_factory=list)
    total_sum: float = 0.0


class CommercialPileService:
    """Build pile order_data from normalized text + pile_prices lookup."""

    def generate_preview(
        self,
        text: str,
        *,
        db_path: str,
        default_grade: str = "B25",
    ) -> CommercialPilePreviewResult:
        normalized = normalize_pile_order_text(text)
        raw_lines = normalized.normalized_lines or [
            part.strip() for part in (text or "").splitlines() if part.strip()
        ]

        parsed_lines = [
            parse_pile_line(line, default_grade=default_grade) for line in raw_lines
        ]
        unparsed_lines = [
            raw_lines[idx]
            for idx, result in enumerate(parsed_lines)
            if not result.parsed
        ]
        merged = merge_pile_lines(parsed_lines, default_grade=default_grade)

        order_data: list[dict[str, Any]] = []
        total_sum = 0.0
        for item in merged:
            unit_price: float | None
            try:
                unit_price = lookup_pile_price(
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
                    "product_kind": "pile",
                    "name": item.mark,
                    "mark": item.mark,
                    "concrete_grade": item.concrete_grade or default_grade,
                    "qty": qty,
                    "unit_price": unit_price,
                    "line_total": line_total if unit_price is not None else None,
                }
            )

        return CommercialPilePreviewResult(
            normalized_text=normalized.normalized_text,
            normalized_lines=normalized.normalized_lines,
            unparsed_lines=unparsed_lines,
            order_data=order_data,
            total_sum=total_sum,
        )
