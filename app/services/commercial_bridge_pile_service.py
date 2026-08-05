from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from core.bridge_pile_line_parser import merge_bridge_pile_lines, parse_bridge_pile_line
from core.bridge_pile_price_db import (
    list_available_grades,
    resolve_default_bridge_pile_grade,
)
from core.bridge_pile_text_normalizer import normalize_bridge_pile_order_text
from core.commercial_pricing import lookup_bridge_pile_price
from core.exceptions import PriceNotFoundError


@dataclass
class CommercialBridgePilePreviewResult:
    normalized_text: str
    normalized_lines: list[str] = field(default_factory=list)
    unparsed_lines: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    order_data: list[dict[str, Any]] = field(default_factory=list)
    total_sum: float = 0.0


class CommercialBridgePileService:
    """Build bridge-pile order_data from text + bridge_pile_prices lookup."""

    def generate_preview(
        self,
        text: str,
        *,
        db_path: str,
        default_grade: str = "B25",
    ) -> CommercialBridgePilePreviewResult:
        normalized = normalize_bridge_pile_order_text(text)
        raw_lines = normalized.normalized_lines or [
            part.strip() for part in (text or "").splitlines() if part.strip()
        ]

        parsed_lines = [
            parse_bridge_pile_line(line, default_grade=default_grade) for line in raw_lines
        ]
        unparsed_lines = [
            raw_lines[idx]
            for idx, result in enumerate(parsed_lines)
            if not result.parsed
        ]
        merged = merge_bridge_pile_lines(parsed_lines, default_grade=default_grade)

        order_data: list[dict[str, Any]] = []
        total_sum = 0.0
        for item in merged:
            available = list_available_grades(item.mark, db_path=db_path)
            preferred = item.concrete_grade or default_grade
            grade = resolve_default_bridge_pile_grade(
                item.mark,
                preferred=preferred if preferred in (available or [preferred]) else None,
                db_path=db_path,
            )
            if grade is None:
                grade = preferred

            unit_price: float | None
            try:
                unit_price = lookup_bridge_pile_price(
                    item.mark,
                    grade,
                    db_path=db_path,
                )
            except PriceNotFoundError:
                unit_price = None

            qty = int(item.qty or 0)
            line_total = (unit_price or 0.0) * qty
            total_sum += line_total
            order_data.append(
                {
                    "product_kind": "bridge_pile",
                    "name": item.mark,
                    "mark": item.mark,
                    "concrete_grade": grade,
                    "available_grades": available,
                    "qty": qty,
                    "unit_price": unit_price,
                    "line_total": line_total if unit_price is not None else None,
                }
            )

        return CommercialBridgePilePreviewResult(
            normalized_text=normalized.normalized_text,
            normalized_lines=normalized.normalized_lines,
            unparsed_lines=unparsed_lines,
            order_data=order_data,
            total_sum=total_sum,
        )
