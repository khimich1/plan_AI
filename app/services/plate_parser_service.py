from __future__ import annotations

import logging
import sqlite3
from pathlib import Path
from typing import Any

from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.repositories.kp_repository import KpRepository
from core.config_and_data import make_plate_name
from core.exceptions import PlateParseError
from core.kp_db import lookup_nomenclature_by_plate_name
from core.plate_line_parser import build_lwh_mm_load_warning, parse_line
from core.plate_text_normalizer import get_wide_plate_lines, normalize_order_text
from core.plate_validation import validate_plate_values

logger = logging.getLogger(__name__)


class PlateParserService:
    def __init__(self) -> None:
        self.kp_repository = KpRepository()
        self.pb_db_path = Path(self.kp_repository.db_path).resolve().parent / "pb.db"

    def parse_plate_text(self, text: str) -> ParseResult:
        if not text or not text.strip():
            raise PlateParseError(
                "Текст заказа пустой. Пожалуйста, введите список плит.\n"
                "Пример: ПБ 78-12-8п 5 шт"
            )

        normalization = normalize_order_text(text)
        processing_text = normalization.normalized_text.strip() or text
        normalized_lines = list(normalization.normalized_lines)
        lines = [line for line in normalized_lines if line.strip()]
        if not lines:
            lines = [line.strip() for line in processing_text.replace("×", "x").splitlines() if line.strip()]
        if not lines:
            raise PlateParseError("Не удалось найти ни одной строки с плитами.\nПроверьте формат ввода.")

        order = PlateOrder()
        diagnostics: list[dict[str, Any]] = []
        unparsed_lines: list[str] = []
        line_contributions: list[list[tuple[float, float, float | None, str]]] = [[] for _ in lines]
        line_plate_load_details: list[dict[tuple, int]] = [{} for _ in lines]
        lwh_mm_assumed_lines: list[str] = []

        def record_contribution(
            line_idx: int,
            *,
            length_m: float,
            width_m: float,
            load_code: float | None,
            length_dm_raw: str,
        ) -> None:
            line_contributions[line_idx].append(
                (round(length_m, 3), round(width_m, 3), load_code, (length_dm_raw or "").strip())
            )

        def add_items(
            *,
            line_idx: int,
            width_m: float,
            length_m: float,
            qty: int,
            load_code: float | None,
            length_dm_raw: str,
        ) -> None:
            width_rounded = round(width_m, 3)
            length_rounded = round(length_m, 3)
            target_name: str | None = None
            target: list[float] | None = None

            if abs(width_rounded - 1.2) < 0.01:
                target_name = "plates_1_2"
                target = order.plates_1_2
            elif abs(width_rounded - 1.08) < 0.07 or 1.02 <= width_rounded <= 1.08:
                target_name = "plates_1_08"
                target = order.plates_1_08
            elif abs(width_rounded - 1.0) < 0.05:
                target_name = "plates_1_0"
                target = order.plates_1_0
            elif 0.46 <= width_rounded <= 0.53:
                target_name = "plates_0_46"
                target = order.plates_0_46
            elif 0.26 <= width_rounded <= 0.32:
                target_name = "plates_0_32"
                target = order.plates_0_32
            elif 0.66 <= width_rounded <= 0.70:
                target_name = "plates_0_70"
                target = order.plates_0_70
            elif 0.70 <= width_rounded <= 0.72:
                target_name = "plates_0_72"
                target = order.plates_0_72
            elif 0.86 <= width_rounded <= 0.92:
                target_name = "plates_0_86"
                target = order.plates_0_86
            elif 0.72 <= width_rounded <= 0.74:
                target_name = "plates_0_74"
                target = order.plates_0_74
            elif 0.48 <= width_rounded <= 0.50:
                target_name = "plates_0_48"
                target = order.plates_0_48
            elif 0.50 <= width_rounded <= 0.53:
                target_name = "plates_0_50"
                target = order.plates_0_50
            elif 0.34 <= width_rounded <= 0.36:
                target_name = "plates_0_34"
                target = order.plates_0_34

            if target is None:
                target_name = "plates_1_2"
                target = order.plates_1_2

            for _ in range(max(0, qty)):
                target.append(length_rounded)
                order.plate_exact_widths[(length_rounded, target_name)] = width_rounded

            if load_code is not None and load_code > 0:
                key = (length_rounded, width_rounded, float(load_code), (length_dm_raw or "").strip())
                order.plate_load_details[key] = order.plate_load_details.get(key, 0) + qty
                order.plate_length_dm_raw[key] = (length_dm_raw or "").strip()
                line_plate_load_details[line_idx][key] = line_plate_load_details[line_idx].get(key, 0) + qty
                record_contribution(
                    line_idx,
                    length_m=length_rounded,
                    width_m=width_rounded,
                    load_code=float(load_code),
                    length_dm_raw=length_dm_raw,
                )
            else:
                record_contribution(
                    line_idx,
                    length_m=length_rounded,
                    width_m=width_rounded,
                    load_code=None,
                    length_dm_raw=length_dm_raw,
                )

        for line_idx, raw in enumerate(lines):
            parsed_line = parse_line(raw)
            diagnostic: dict[str, Any] = {
                "raw_input": raw,
                "parse_stage": parsed_line.stage,
                "recognized_by": "parser",
            }
            if not parsed_line.parsed:
                diagnostic["validation_status"] = "failed"
                diagnostic["reason_code"] = parsed_line.reason_code or "pattern_not_matched"
                diagnostic["rejection_reason"] = parsed_line.reason_text or "строка не распознана"
                diagnostics.append(diagnostic)
                unparsed_lines.append(f"{raw} (пропущено: {diagnostic['rejection_reason']})")
                continue

            validation = validate_plate_values(parsed_line.width_m, parsed_line.length_m, parsed_line.qty)
            if not validation.ok:
                diagnostic["validation_status"] = "failed"
                diagnostic["reason_code"] = validation.reason_code
                diagnostic["rejection_reason"] = validation.reason_text
                diagnostics.append(diagnostic)
                unparsed_lines.append(f"{raw} (пропущено: {validation.reason_text})")
                continue

            add_items(
                line_idx=line_idx,
                width_m=parsed_line.width_m,
                length_m=parsed_line.length_m,
                qty=parsed_line.qty,
                load_code=parsed_line.load_code,
                length_dm_raw=parsed_line.length_dm_raw,
            )
            diagnostic["validation_status"] = "ok"
            diagnostic["normalized_input"] = raw
            if parsed_line.load_assumed:
                diagnostic["load_assumed"] = True
                diagnostic["load_warning_code"] = "lwh_mm_default_load"
                lwh_mm_assumed_lines.append(raw)
            diagnostics.append(diagnostic)

        order.recompute_totals()
        self._fill_nomenclature_cache(order)

        warnings = list(normalization.warnings)
        if lwh_mm_assumed_lines:
            warnings.append(build_lwh_mm_load_warning(lwh_mm_assumed_lines))
        if unparsed_lines:
            warnings.append(f"Не удалось распознать строк: {len(unparsed_lines)}")

        return ParseResult(
            order=order,
            normalized_text=processing_text,
            normalized_lines=lines,
            unparsed_lines=unparsed_lines,
            diagnostics=diagnostics,
            wide_plate_lines=get_wide_plate_lines(processing_text),
            warnings=warnings,
            line_contributions=line_contributions,
            line_plate_load_details=line_plate_load_details,
        )

    def _fill_nomenclature_cache(self, order: PlateOrder) -> None:
        if not self.pb_db_path.exists() or not order.plate_load_details:
            return

        with sqlite3.connect(str(self.pb_db_path)) as conn:
            cursor = conn.cursor()
            for key in order.plate_load_details:
                length_m, width_m, load_code, length_dm_raw = key
                plate_name = make_plate_name(
                    length_m,
                    width_m,
                    load_code=int(float(load_code)),
                    length_dm_raw=length_dm_raw,
                )
                canonical_name, nomenclature_id, _match_type = lookup_nomenclature_by_plate_name(plate_name, cursor)
                order.nomenclature_cache[key] = {
                    "canonical_name": canonical_name,
                    "nomenclature_id": nomenclature_id,
                }

