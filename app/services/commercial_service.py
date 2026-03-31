from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.repositories.manager_repository import ManagerRepository
from app.services.file_generation_service import FileGenerationService
from app.services.optimization_service import OptimizationService
from app.services.plate_parser_service import PlateParserService
from core import config_and_data as cfg
from core.kp_db import enrich_order_data_with_nomenclature
from core.kp_plate_weight import resolve_kp_line_weight_kg
from viz_modules.price_utils import load_price_table_from_xlsx
from viz_modules.procurement import build_component_breakdown, build_price_rows, build_procurement_items


@dataclass
class CommercialPreviewResult:
    parse_result: ParseResult
    optimization_context: OptimizationContext
    order_data: list[dict[str, Any]]
    price_rows: list
    breakdown_tables: list[dict[str, Any]]
    total_sum: float


class CommercialService:
    def __init__(self) -> None:
        self.manager_repository = ManagerRepository()
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
    ) -> CommercialPreviewResult:
        current_parse_result = parse_result or self.parse(text or "")
        order = current_parse_result.order
        optimization_context = self.optimization_service.optimize(order)

        with self.file_generation_service._legacy_order_context(order):
            with self.optimization_service.legacy_runtime(optimization_context):
                price_table = load_price_table_from_xlsx(str(cfg.PRICE_XLSX_PATH))
                price_rows, total_sum = build_price_rows(price_table, reinforcement_code=8)
                breakdown_tables = build_component_breakdown(price_table, price_rows)
                procurement_items = build_procurement_items()

        order_data = self._build_order_data(procurement_items, price_rows, order)
        order_data = enrich_order_data_with_nomenclature(order_data)
        return CommercialPreviewResult(
            parse_result=current_parse_result,
            optimization_context=optimization_context,
            order_data=order_data,
            price_rows=price_rows,
            breakdown_tables=breakdown_tables,
            total_sum=float(total_sum),
        )

    def _build_order_data(
        self,
        procurement_items: list[dict[str, Any]],
        price_rows: list,
        order: PlateOrder,
    ) -> list[dict[str, Any]]:
        order_data: list[dict[str, Any]] = []
        cache_by_key = order.nomenclature_cache
        for item in procurement_items:
            length_m = float(item["length"])
            width_m = float(item["width"])
            qty = int(item["qty"])
            load_code = item.get("load_code")
            if load_code is None:
                load_code = cfg.get_load_code_for_plate(length_m, width_m, default=(6 if width_m < 1.0 else 8))

            matching_row = None
            for row in price_rows:
                if len(row) < 8:
                    continue
                row_name = row[1]
                parsed_length, parsed_width = cfg.parse_name_to_sizes(row_name)
                if parsed_length is None or parsed_width is None:
                    continue
                parsed_load = cfg.parse_load_code_from_name(row_name)
                if (
                    abs(parsed_length - length_m) < 0.01
                    and abs(parsed_width - width_m) < 0.01
                    and cfg.load_code_for_price_match(parsed_load) == cfg.load_code_for_price_match(load_code)
                ):
                    matching_row = row
                    break

            item_ldr = item.get("length_dm_raw") or ""
            name = cfg.make_plate_name(length_m, width_m, load_code=load_code, length_dm_raw=item_ldr or None)
            unit_price = 0.0
            length_dm_raw = item_ldr
            if matching_row:
                name = matching_row[1]
                try:
                    unit_price = float(str(matching_row[7]).replace(" ", "").replace(",", "."))
                except (TypeError, ValueError):
                    unit_price = 0.0
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
                "load_class": (cfg.normalize_load_code(load_code) or 8) * 100,
                "unit_price": unit_price,
                "weight": total_weight_kg,
            }
            if nomenclature_id is not None:
                entry["nomenclature_id"] = nomenclature_id
            order_data.append(entry)
        return order_data

