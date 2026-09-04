from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.repositories.kp_repository import KpRepository
from app.repositories.manager_repository import ManagerRepository
from app.services.file_generation_service import FileGenerationService
from app.services.optimization_service import OptimizationService
from app.services.plate_parser_service import PlateParserService
from core.commercial_offer_xlsx import DB_PATH
from core.config_and_data import (
    load_code_for_price_match,
    make_plate_name,
    parse_load_code_from_name,
    parse_name_to_sizes,
)
from core.domain.plate_order import normalize_load_code
from core.optimization.layout_runtime_snapshot import _make_get_load_code_for_plate
from core.project_paths import PRICE_XLSX_PATH
from core.plate_order_context import PlateOrderContext
from core.kp_db_nomenclature import enrich_order_data_with_nomenclature
from core.kp_plate_weight import resolve_kp_line_weight_kg
from core.ports.visualization import (
    build_component_breakdown,
    build_price_rows,
    build_procurement_items,
    load_price_table_from_xlsx,
)
from core.invalid_width_lines import build_invalid_width_lines
from core.unpriced_plate_replacements import build_unpriced_plate_lines


@dataclass
class CommercialPreviewResult:
    parse_result: ParseResult
    optimization_context: OptimizationContext
    order_data: list[dict[str, Any]]
    price_rows: list
    breakdown_tables: list[dict[str, Any]]
    total_sum: float
    unpriced_plate_lines: list[dict[str, Any]] = field(default_factory=list)
    invalid_width_lines: list[dict[str, Any]] = field(default_factory=list)


class CommercialService:
    def __init__(self) -> None:
        self.manager_repository = ManagerRepository()
        self.kp_repository = KpRepository()
        self.parser_service = PlateParserService()
        self.optimization_service = OptimizationService()
        self.file_generation_service = FileGenerationService()

    def list_managers(self) -> list[dict]:
        return self.manager_repository.list_managers()

    def parse(self, text: str) -> ParseResult:
        return self.parser_service.parse_plate_text(text)

    def generate_preview(
        self,
        *,
        text: str | None = None,
        parse_result: ParseResult | None = None,
        plate_order_ctx: PlateOrderContext,
    ) -> CommercialPreviewResult:
        current_parse_result = parse_result or self.parse(text or "")
        order = current_parse_result.order
        ctx = plate_order_ctx
        optimization_context = self.optimization_service.optimize(
            order,
            plate_order_ctx=ctx,
        )
        ctx.load_optimization_snapshot(
            optimization_result=optimization_context.optimization_result,
            plan_by_load=optimization_context.plan_by_load,
            load_to_reinforcement_map=optimization_context.load_to_reinforcement_map,
        )
        with ctx.bound():
            price_table = load_price_table_from_xlsx(str(PRICE_XLSX_PATH))
            load_kwargs = {"plate_load_details": order.plate_load_details}
            price_rows, total_sum = build_price_rows(
                price_table,
                reinforcement_code=8,
                **load_kwargs,
            )
            breakdown_tables = build_component_breakdown(
                price_table,
                price_rows,
                **load_kwargs,
            )
            procurement_items = build_procurement_items(**load_kwargs)

        order_data = self._build_order_data(
            procurement_items,
            price_rows,
            order,
            current_parse_result,
        )
        order_data = enrich_order_data_with_nomenclature(order_data)
        unpriced_plate_lines = build_unpriced_plate_lines(
            order_data,
            db_path=str(DB_PATH),
            normalized_lines=list(current_parse_result.normalized_lines),
        )
        skip_wide = [
            str(item[0]).strip()
            for item in current_parse_result.wide_plate_lines
            if item
        ]
        invalid_width_lines = build_invalid_width_lines(
            order_data,
            db_path=str(DB_PATH),
            normalized_lines=list(current_parse_result.normalized_lines),
            skip_wide_lines=skip_wide,
        )
        return CommercialPreviewResult(
            parse_result=current_parse_result,
            optimization_context=optimization_context,
            order_data=order_data,
            price_rows=price_rows,
            breakdown_tables=breakdown_tables,
            total_sum=float(total_sum),
            unpriced_plate_lines=unpriced_plate_lines,
            invalid_width_lines=invalid_width_lines,
        )

    def _build_order_data(
        self,
        procurement_items: list[dict[str, Any]],
        price_rows: list,
        order: PlateOrder,
        parse_result: ParseResult,
    ) -> list[dict[str, Any]]:
        order_data_with_order: list[tuple[int, dict[str, Any]]] = []
        cache_by_key = order.nomenclature_cache
        order_sequence = self._build_order_sequence_map(parse_result)
        resolve_load_code = _make_get_load_code_for_plate(order.plate_load_details)
        for item in procurement_items:
            length_m = float(item["length"])
            width_m = float(item["width"])
            qty = int(item["qty"])
            load_code = item.get("load_code")
            if load_code is None:
                load_code = resolve_load_code(
                    length_m,
                    width_m,
                    default=(6 if width_m < 1.0 else 8),
                )

            matching_row = None
            for row in price_rows:
                if len(row) < 8:
                    continue
                row_name = row[1]
                parsed_length, parsed_width = parse_name_to_sizes(row_name)
                if parsed_length is None or parsed_width is None:
                    continue
                parsed_load = parse_load_code_from_name(row_name)
                if (
                    abs(parsed_length - length_m) < 0.01
                    and abs(parsed_width - width_m) < 0.01
                    and load_code_for_price_match(parsed_load) == load_code_for_price_match(load_code)
                ):
                    matching_row = row
                    break

            item_ldr = item.get("length_dm_raw") or ""
            name = make_plate_name(length_m, width_m, load_code=load_code, length_dm_raw=item_ldr or None)
            unit_price: float | None = None
            length_dm_raw = item_ldr
            if matching_row:
                name = matching_row[1]
                try:
                    parsed_price = float(str(matching_row[7]).replace(" ", "").replace(",", "."))
                    unit_price = parsed_price if parsed_price > 0 else None
                except (TypeError, ValueError):
                    unit_price = None
                match = re.search(r"ПБ\s+([\d,]+)-", name)
                if match:
                    length_dm_raw = match.group(1).strip()

            cache_key = (round(length_m, 3), round(width_m, 3), float(load_code), length_dm_raw)
            cache_value = cache_by_key.get(cache_key) or cache_by_key.get(
                (round(length_m, 3), round(width_m, 3), float(load_code), item_ldr)
            )
            if cache_value:
                if cache_value.get("canonical_name"):
                    name = cache_value["canonical_name"]
                nomenclature_id = cache_value.get("nomenclature_id")
            else:
                nomenclature_id = None

            _, total_weight_kg = resolve_kp_line_weight_kg(
                {
                    "length_m": length_m,
                    "width_m": width_m,
                    "qty": qty,
                }
            )
            entry = {
                "name": name,
                "length_m": length_m,
                "length_dm_raw": length_dm_raw or item_ldr,
                "width_m": width_m,
                "qty": qty,
                "load_class": (normalize_load_code(load_code) or 8) * 100,
                "unit_price": unit_price,
                "weight": total_weight_kg,
            }
            if nomenclature_id is not None:
                entry["nomenclature_id"] = nomenclature_id
            line_order = self._resolve_line_order(
                order_sequence,
                length_m=length_m,
                width_m=width_m,
                load_code=load_code,
                length_dm_raw=(length_dm_raw or item_ldr),
                fallback_raw=item_ldr,
            )
            order_data_with_order.append((line_order, entry))

        order_data_with_order.sort(key=lambda item: item[0])
        return [entry for _order, entry in order_data_with_order]

    @staticmethod
    def _normalize_sequence_key(
        *,
        length_m: float,
        width_m: float,
        load_code: float | int,
        length_dm_raw: str,
    ) -> tuple[float, float, float, str]:
        return (
            round(float(length_m), 3),
            round(float(width_m), 3),
            float(load_code),
            (length_dm_raw or "").strip(),
        )

    def _build_order_sequence_map(self, parse_result: ParseResult) -> dict[tuple[float, float, float, str], int]:
        sequence: dict[tuple[float, float, float, str], int] = {}
        for line_idx, line_items in enumerate(parse_result.line_plate_load_details):
            for key in line_items.keys():
                if len(key) < 4:
                    continue
                sequence_key = self._normalize_sequence_key(
                    length_m=float(key[0]),
                    width_m=float(key[1]),
                    load_code=float(key[2]),
                    length_dm_raw=str(key[3]),
                )
                if sequence_key not in sequence:
                    sequence[sequence_key] = line_idx
        return sequence

    def _resolve_line_order(
        self,
        order_sequence: dict[tuple[float, float, float, str], int],
        *,
        length_m: float,
        width_m: float,
        load_code: float | int,
        length_dm_raw: str,
        fallback_raw: str,
    ) -> int:
        default_order = 10_000
        candidate_keys = [
            self._normalize_sequence_key(
                length_m=length_m,
                width_m=width_m,
                load_code=load_code,
                length_dm_raw=length_dm_raw,
            ),
            self._normalize_sequence_key(
                length_m=length_m,
                width_m=width_m,
                load_code=load_code,
                length_dm_raw=fallback_raw,
            ),
            self._normalize_sequence_key(
                length_m=length_m,
                width_m=width_m,
                load_code=load_code,
                length_dm_raw="",
            ),
        ]
        for candidate in candidate_keys:
            if candidate in order_sequence:
                return order_sequence[candidate]
        return default_order

    def save_offer(
        self,
        *,
        creation_date: str,
        order_data: list[dict[str, Any]],
        xlsx_path: str | None = None,
        customer_name: str | None = None,
        manager_name: str | None = None,
        discount_percent: float = 0.0,
        delivery_conditions: str | None = None,
        payment_conditions: str | None = None,
        execution_terms: str | None = None,
        status: str = "в работе",
    ) -> int:
        """Сохранить КП в БД (тонкая обёртка над ``KpRepository`` для bot/web)."""
        return self.kp_repository.save_offer(
            creation_date=creation_date,
            order_data=order_data,
            xlsx_path=xlsx_path,
            customer_name=customer_name,
            manager_name=manager_name,
            discount_percent=discount_percent,
            delivery_conditions=delivery_conditions or "",
            payment_conditions=payment_conditions or "",
            execution_terms=execution_terms or "",
            status=status,
        )

    def get_offer(self, kp_id: int) -> dict | None:
        return self.kp_repository.get_offer(kp_id)

