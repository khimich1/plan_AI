from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from core.commercial_pricing import lookup_step_price
from core.exceptions import PriceNotFoundError
from core.step_line_parser import merge_step_lines, parse_step_line
from core.step_text_normalizer import normalize_step_order_text


@dataclass
class CommercialStepPreviewResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)
    unparsed_lines: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    order_data: list[dict[str, Any]] = field(default_factory=list)
    total_sum: float = 0.0


class CommercialStepService:
    """Build step order_data from normalized text + step_prices lookup (no grade)."""

    def generate_preview(
        self,
        text: str,
        *,
        db_path: str,
    ) -> CommercialStepPreviewResult:
        normalized = normalize_step_order_text(text)
        raw_lines = normalized.normalized_lines or [
            part.strip() for part in (text or "").splitlines() if part.strip()
        ]

        parsed_lines = [parse_step_line(line) for line in raw_lines]
        unparsed_lines = [
            raw_lines[idx]
            for idx, result in enumerate(parsed_lines)
            if not result.parsed
        ]
        merged = merge_step_lines(parsed_lines)

        order_data: list[dict[str, Any]] = []
        total_sum = 0.0
        for item in merged:
            unit_price: float | None
            try:
                unit_price = lookup_step_price(item.mark, db_path=db_path)
            except PriceNotFoundError:
                unit_price = None

            qty = int(item.qty or 0)
            line_total = (unit_price or 0.0) * qty
            total_sum += line_total
            order_data.append(
                {
                    "product_kind": "step",
                    "name": item.mark,
                    "mark": item.mark,
                    "qty": qty,
                    "unit_price": unit_price,
                    "line_total": line_total if unit_price is not None else None,
                }
            )

        return CommercialStepPreviewResult(
            normalized_text=normalized.normalized_text,
            normalized_lines=normalized.normalized_lines,
            unparsed_lines=unparsed_lines,
            order_data=order_data,
            total_sum=total_sum,
        )
