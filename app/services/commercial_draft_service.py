from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable

from app.core.settings import get_settings
from app.schemas.commercial import WizardStepId
from app.services.commercial_service import CommercialService
from core.ocr_gpt import recognize_text_smart

_ALLOWED_OCR_IMAGE_SUFFIXES = frozenset({".jpg", ".jpeg", ".png", ".webp", ".gif", ".pdf"})


def safe_ocr_temp_suffix(image_filename: str | None) -> str:
    """Use only a whitelisted extension from the upload basename; never user-controlled path segments."""
    raw = (image_filename or "").strip()
    if not raw:
        return ".jpg"
    base = Path(raw).name
    if ".." in base or "/" in base or "\\" in base:
        return ".jpg"
    suffix = Path(base).suffix.lower()
    if suffix in _ALLOWED_OCR_IMAGE_SUFFIXES:
        return suffix
    return ".jpg"


# Backward-compatible alias for tests importing from commercial_workflow_service.
_safe_ocr_temp_suffix = safe_ocr_temp_suffix


_OCR_VERIFY_SKIPPED_WARNING = (
    "Распознавание выполнено без второй проверки (OCR_MAX_API_CALLS=1). "
    "Проверьте марки и количество вручную."
)


class CommercialDraftService:
    """Draft ingest helpers: OCR/text source resolution and preview metadata assembly."""

    def __init__(self, *, commercial_service: CommercialService | None = None) -> None:
        self.settings = get_settings()
        self.commercial_service = commercial_service or CommercialService()

    def _ocr_quality_warnings(self, result: dict[str, Any]) -> list[str]:
        """Q4: при OCR_MAX_API_CALLS=1 предупреждение о сомнительном Extract."""
        if int(self.settings.ocr_max_api_calls or 2) > 1:
            return []

        plates = list(result.get("plates") or [])
        min_confidence = float(self.settings.ocr_verify_auto_min_confidence or 0.92)
        has_low_confidence = any(self._plate_confidence(plate) < min_confidence for plate in plates)
        has_parser_rejected = any(
            "parser_rejected" in list(plate.get("issues") or [])
            for plate in plates
        )
        if has_low_confidence or has_parser_rejected:
            return [_OCR_VERIFY_SKIPPED_WARNING]
        return []

    @staticmethod
    def _plate_confidence(plate: dict[str, Any]) -> float:
        try:
            return float(plate.get("confidence", 0.95))
        except (TypeError, ValueError):
            return 0.95

    def _map_ocr_result_metadata(
        self,
        result: dict[str, Any],
        *,
        recognition_mode: str,
    ) -> dict[str, Any]:
        return {
            "ocr_recognition_mode": recognition_mode,
            "ocr_cost_usd": float(result.get("cost_usd", 0.0) or 0.0),
            "ocr_cost_rub": float(result.get("ocr_cost_rub", 0.0) or 0.0),
            "ocr_api_calls": int(result.get("ocr_api_calls", 1) or 1),
            "ocr_plates": list(result.get("plates") or []),
            "ocr_draft_plates": list(result.get("draft_plates") or []),
            "ocr_corrections": list(result.get("corrections") or []),
            "ocr_verify_applied": bool(result.get("verify_applied")),
            "ocr_verify_failed": bool(result.get("verify_failed")),
            "ocr_method": str(result.get("ocr_method") or result.get("method") or "GPT-4o"),
            "ocr_row_count_on_image": result.get("row_count_on_image"),
            "ocr_verify_skipped_reason": result.get("ocr_verify_skipped_reason"),
            "ocr_verify_applied_reason": result.get("ocr_verify_applied_reason"),
            "ocr_warnings": self._ocr_quality_warnings(result),
        }

    async def extract_text_from_image(
        self,
        *,
        image_bytes: bytes,
        image_filename: str | None,
        product_type: str = "plates",
    ) -> tuple[str, dict[str, Any]]:
        suffix = safe_ocr_temp_suffix(image_filename)
        with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(image_bytes)
            tmp_path = Path(tmp_file.name)

        recognition_mode = (self.settings.ocr_recognition_mode or "full_gpt").strip().lower()
        if recognition_mode not in {"full_gpt", "hybrid"}:
            recognition_mode = "full_gpt"

        normalized_product_type = (product_type or "plates").strip().lower()
        if normalized_product_type not in {"plates", "piles", "steps", "marches", "bridge_piles", "fbs"}:
            normalized_product_type = "plates"

        try:
            result = await recognize_text_smart(
                str(tmp_path),
                force_gpt=(recognition_mode == "full_gpt"),
                show_cost=True,
                mode=recognition_mode,
                product_type=normalized_product_type,  # type: ignore[arg-type]
            )
        finally:
            tmp_path.unlink(missing_ok=True)

        if not result:
            raise ValueError("Не удалось распознать текст на изображении.")

        recognized_text = str(result.get("text", "")).strip()
        if not recognized_text:
            raise ValueError("Не удалось распознать текст на изображении.")
        return recognized_text, self._map_ocr_result_metadata(
            result,
            recognition_mode=recognition_mode,
        )

    async def resolve_source_input(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        product_type: str = "plates",
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            raise ValueError("Нужно передать текст или изображение для распознавания.")
        if image_bytes and not text_value:
            recognized_text, source_metadata = await self.extract_text_from_image(
                image_bytes=image_bytes,
                image_filename=image_filename,
                product_type=product_type,
            )
            return (
                {
                    "source_type": "image",
                    "original_text": text_value,
                    "ocr_text": recognized_text,
                    "input_text": recognized_text,
                    "filename": image_filename or "",
                    "batch": {
                        "source_type": "image",
                        "original_text": text_value,
                        "normalized_text": recognized_text,
                        "ocr_text": recognized_text,
                        "filename": image_filename or "",
                    },
                },
                source_metadata,
            )
        return (
            {
                "source_type": "text",
                "original_text": text_value,
                "ocr_text": "",
                "input_text": text_value,
                "filename": "",
                "batch": {
                    "source_type": "text",
                    "original_text": text_value,
                    "normalized_text": text_value,
                    "ocr_text": "",
                    "filename": "",
                },
            },
            {},
        )

    def build_preview_metadata(
        self,
        *,
        preview: Any,
        base_metadata: dict[str, Any],
        source_type: str,
        original_text: str,
        ocr_text: str,
        input_text: str,
        last_source_filename: str,
        plate_batches: list[dict[str, Any]],
        wide_plates_resolved: bool,
        source_metadata: dict[str, Any],
        owner_user_id: int | None = None,
    ) -> dict[str, Any]:
        source_metadata_payload = dict(source_metadata)
        ocr_warnings = list(source_metadata_payload.pop("ocr_warnings", []) or [])
        warnings = list(preview.parse_result.warnings)
        for warning in ocr_warnings:
            if warning not in warnings:
                warnings.append(warning)

        metadata = dict(base_metadata)
        metadata.update(
            {
                "product_type": metadata.get("product_type") or "plates",
                "source_type": source_type,
                "original_text": original_text,
                "ocr_text": ocr_text,
                "input_text": input_text,
                "accumulated_text": input_text,
                "warnings": warnings,
                "unparsed_lines": list(preview.parse_result.unparsed_lines),
                "normalized_text": preview.parse_result.normalized_text,
                "normalized_lines": list(preview.parse_result.normalized_lines),
                "wide_plate_lines": self.serialize_wide_plate_lines(preview.parse_result.wide_plate_lines),
                "dobor_pairs": self.serialize_dobor_pairs(preview.parse_result.dobor_pairs),
                "diagnostics": list(preview.parse_result.diagnostics),
                "breakdown_tables": list(preview.breakdown_tables),
                "price_rows_count": len(preview.price_rows),
                "breakdown_tables_count": len(preview.breakdown_tables),
                "total_sum": preview.total_sum,
                "plate_batches": plate_batches,
                "wide_plates_resolved": wide_plates_resolved,
                "last_source_filename": last_source_filename,
                "current_step": WizardStepId.plates.value,
                "saved_offer": None,
                "generated_files": [],
                "current_save_mode": None,
                "execution_terms": "",
                **source_metadata_payload,
            }
        )
        metadata.setdefault("manager_id", None)
        metadata.setdefault("manager_name", "")
        metadata.setdefault("manager_phone", "")
        metadata.setdefault("manager_email", "")
        metadata.setdefault("client_name", "")
        metadata.setdefault("discount_percent", 0.0)
        metadata.setdefault("conditions_mode", "standard")
        metadata.setdefault("delivery_conditions", "")
        metadata.setdefault("payment_conditions", "")
        metadata.setdefault("logistics_cost", 0.0)
        if owner_user_id is not None:
            metadata["owner_user_id"] = int(owner_user_id)
        return metadata

    def build_pile_preview_metadata(
        self,
        *,
        preview: Any,
        base_metadata: dict[str, Any],
        source_type: str,
        original_text: str,
        ocr_text: str,
        input_text: str,
        last_source_filename: str,
        pile_batches: list[dict[str, Any]],
        source_metadata: dict[str, Any],
        owner_user_id: int | None = None,
    ) -> dict[str, Any]:
        source_metadata_payload = dict(source_metadata)
        ocr_warnings = list(source_metadata_payload.pop("ocr_warnings", []) or [])
        warnings = list(preview.warnings)
        for warning in ocr_warnings:
            if warning not in warnings:
                warnings.append(warning)

        metadata = dict(base_metadata)
        metadata.update(
            {
                "product_type": "piles",
                "source_type": source_type,
                "original_text": original_text,
                "ocr_text": ocr_text,
                "input_text": input_text,
                "accumulated_text": input_text,
                "warnings": warnings,
                "unparsed_lines": list(preview.unparsed_lines),
                "normalized_text": preview.normalized_text,
                "normalized_lines": list(preview.normalized_lines),
                "wide_plate_lines": [],
                "diagnostics": [],
                "breakdown_tables": [],
                "price_rows_count": len(preview.order_data),
                "breakdown_tables_count": 0,
                "total_sum": preview.total_sum,
                "pile_batches": pile_batches,
                "wide_plates_resolved": True,
                "last_source_filename": last_source_filename,
                "current_step": WizardStepId.piles.value,
                "saved_offer": None,
                "generated_files": [],
                "current_save_mode": None,
                "execution_terms": "",
                **source_metadata_payload,
            }
        )
        metadata.setdefault("manager_id", None)
        metadata.setdefault("manager_name", "")
        metadata.setdefault("manager_phone", "")
        metadata.setdefault("manager_email", "")
        metadata.setdefault("client_name", "")
        metadata.setdefault("discount_percent", 0.0)
        metadata.setdefault("conditions_mode", "standard")
        metadata.setdefault("delivery_conditions", "")
        metadata.setdefault("payment_conditions", "")
        metadata.setdefault("logistics_cost", 0.0)
        metadata.setdefault("default_concrete_grade", "B25")
        if owner_user_id is not None:
            metadata["owner_user_id"] = int(owner_user_id)
        return metadata

    def build_march_preview_metadata(
        self,
        *,
        preview: Any,
        base_metadata: dict[str, Any],
        source_type: str,
        original_text: str,
        ocr_text: str,
        input_text: str,
        last_source_filename: str,
        march_batches: list[dict[str, Any]],
        source_metadata: dict[str, Any],
        owner_user_id: int | None = None,
    ) -> dict[str, Any]:
        source_metadata_payload = dict(source_metadata)
        ocr_warnings = list(source_metadata_payload.pop("ocr_warnings", []) or [])
        warnings = list(preview.warnings)
        for warning in ocr_warnings:
            if warning not in warnings:
                warnings.append(warning)

        metadata = dict(base_metadata)
        metadata.update(
            {
                "product_type": "marches",
                "source_type": source_type,
                "original_text": original_text,
                "ocr_text": ocr_text,
                "input_text": input_text,
                "accumulated_text": input_text,
                "warnings": warnings,
                "unparsed_lines": list(preview.unparsed_lines),
                "normalized_text": preview.normalized_text,
                "normalized_lines": list(preview.normalized_lines),
                "wide_plate_lines": [],
                "diagnostics": [],
                "breakdown_tables": [],
                "price_rows_count": len(preview.order_data),
                "breakdown_tables_count": 0,
                "total_sum": preview.total_sum,
                "march_batches": march_batches,
                "wide_plates_resolved": True,
                "last_source_filename": last_source_filename,
                "current_step": WizardStepId.marches.value,
                "saved_offer": None,
                "generated_files": [],
                "current_save_mode": None,
                "execution_terms": "",
                **source_metadata_payload,
            }
        )
        metadata.setdefault("manager_id", None)
        metadata.setdefault("manager_name", "")
        metadata.setdefault("manager_phone", "")
        metadata.setdefault("manager_email", "")
        metadata.setdefault("client_name", "")
        metadata.setdefault("discount_percent", 0.0)
        metadata.setdefault("conditions_mode", "standard")
        metadata.setdefault("delivery_conditions", "")
        metadata.setdefault("payment_conditions", "")
        metadata.setdefault("logistics_cost", 0.0)
        metadata.setdefault("default_concrete_grade", "B25")
        if owner_user_id is not None:
            metadata["owner_user_id"] = int(owner_user_id)
        return metadata

    def build_bridge_pile_preview_metadata(
        self,
        *,
        preview: Any,
        base_metadata: dict[str, Any],
        source_type: str,
        original_text: str,
        ocr_text: str,
        input_text: str,
        last_source_filename: str,
        bridge_pile_batches: list[dict[str, Any]],
        source_metadata: dict[str, Any],
        owner_user_id: int | None = None,
    ) -> dict[str, Any]:
        source_metadata_payload = dict(source_metadata)
        ocr_warnings = list(source_metadata_payload.pop("ocr_warnings", []) or [])
        warnings = list(preview.warnings)
        for warning in ocr_warnings:
            if warning not in warnings:
                warnings.append(warning)

        metadata = dict(base_metadata)
        metadata.update(
            {
                "product_type": "bridge_piles",
                "source_type": source_type,
                "original_text": original_text,
                "ocr_text": ocr_text,
                "input_text": input_text,
                "accumulated_text": input_text,
                "warnings": warnings,
                "unparsed_lines": list(preview.unparsed_lines),
                "normalized_text": preview.normalized_text,
                "normalized_lines": list(preview.normalized_lines),
                "wide_plate_lines": [],
                "diagnostics": [],
                "breakdown_tables": [],
                "price_rows_count": len(preview.order_data),
                "breakdown_tables_count": 0,
                "total_sum": preview.total_sum,
                "bridge_pile_batches": bridge_pile_batches,
                "wide_plates_resolved": True,
                "last_source_filename": last_source_filename,
                "current_step": WizardStepId.bridge_piles.value,
                "saved_offer": None,
                "generated_files": [],
                "current_save_mode": None,
                "execution_terms": "",
                **source_metadata_payload,
            }
        )
        metadata.setdefault("manager_id", None)
        metadata.setdefault("manager_name", "")
        metadata.setdefault("manager_phone", "")
        metadata.setdefault("manager_email", "")
        metadata.setdefault("client_name", "")
        metadata.setdefault("discount_percent", 0.0)
        metadata.setdefault("conditions_mode", "standard")
        metadata.setdefault("delivery_conditions", "")
        metadata.setdefault("payment_conditions", "")
        metadata.setdefault("logistics_cost", 0.0)
        metadata.setdefault("default_concrete_grade", "B25")
        if owner_user_id is not None:
            metadata["owner_user_id"] = int(owner_user_id)
        return metadata

    def build_fbs_preview_metadata(
        self,
        *,
        preview: Any,
        base_metadata: dict[str, Any],
        source_type: str,
        original_text: str,
        ocr_text: str,
        input_text: str,
        last_source_filename: str,
        fbs_batches: list[dict[str, Any]],
        source_metadata: dict[str, Any],
        owner_user_id: int | None = None,
    ) -> dict[str, Any]:
        source_metadata_payload = dict(source_metadata)
        ocr_warnings = list(source_metadata_payload.pop("ocr_warnings", []) or [])
        warnings = list(preview.warnings)
        for warning in ocr_warnings:
            if warning not in warnings:
                warnings.append(warning)

        metadata = dict(base_metadata)
        metadata.update(
            {
                "product_type": "fbs",
                "source_type": source_type,
                "original_text": original_text,
                "ocr_text": ocr_text,
                "input_text": input_text,
                "accumulated_text": input_text,
                "warnings": warnings,
                "unparsed_lines": list(preview.unparsed_lines),
                "normalized_text": preview.normalized_text,
                "normalized_lines": list(preview.normalized_lines),
                "wide_plate_lines": [],
                "diagnostics": [],
                "breakdown_tables": [],
                "price_rows_count": len(preview.order_data),
                "breakdown_tables_count": 0,
                "total_sum": preview.total_sum,
                "fbs_batches": fbs_batches,
                "wide_plates_resolved": True,
                "last_source_filename": last_source_filename,
                "current_step": WizardStepId.fbs.value,
                "saved_offer": None,
                "generated_files": [],
                "current_save_mode": None,
                "execution_terms": "",
                **source_metadata_payload,
            }
        )
        metadata.setdefault("manager_id", None)
        metadata.setdefault("manager_name", "")
        metadata.setdefault("manager_phone", "")
        metadata.setdefault("manager_email", "")
        metadata.setdefault("client_name", "")
        metadata.setdefault("discount_percent", 0.0)
        metadata.setdefault("conditions_mode", "standard")
        metadata.setdefault("delivery_conditions", "")
        metadata.setdefault("payment_conditions", "")
        metadata.setdefault("logistics_cost", 0.0)
        metadata.setdefault("default_concrete_grade", "B25")
        if owner_user_id is not None:
            metadata["owner_user_id"] = int(owner_user_id)
        return metadata

    def build_step_preview_metadata(
        self,
        *,
        preview: Any,
        base_metadata: dict[str, Any],
        source_type: str,
        original_text: str,
        ocr_text: str,
        input_text: str,
        last_source_filename: str,
        step_batches: list[dict[str, Any]],
        source_metadata: dict[str, Any],
        owner_user_id: int | None = None,
    ) -> dict[str, Any]:
        source_metadata_payload = dict(source_metadata)
        ocr_warnings = list(source_metadata_payload.pop("ocr_warnings", []) or [])
        warnings = list(preview.warnings)
        for warning in ocr_warnings:
            if warning not in warnings:
                warnings.append(warning)

        metadata = dict(base_metadata)
        metadata.update(
            {
                "product_type": "steps",
                "source_type": source_type,
                "original_text": original_text,
                "ocr_text": ocr_text,
                "input_text": input_text,
                "accumulated_text": input_text,
                "warnings": warnings,
                "unparsed_lines": list(preview.unparsed_lines),
                "normalized_text": preview.normalized_text,
                "normalized_lines": list(preview.normalized_lines),
                "wide_plate_lines": [],
                "diagnostics": [],
                "breakdown_tables": [],
                "price_rows_count": len(preview.order_data),
                "breakdown_tables_count": 0,
                "total_sum": preview.total_sum,
                "step_batches": step_batches,
                "wide_plates_resolved": True,
                "last_source_filename": last_source_filename,
                "current_step": WizardStepId.steps.value,
                "saved_offer": None,
                "generated_files": [],
                "current_save_mode": None,
                "execution_terms": "",
                **source_metadata_payload,
            }
        )
        metadata.setdefault("manager_id", None)
        metadata.setdefault("manager_name", "")
        metadata.setdefault("manager_phone", "")
        metadata.setdefault("manager_email", "")
        metadata.setdefault("client_name", "")
        metadata.setdefault("discount_percent", 0.0)
        metadata.setdefault("conditions_mode", "standard")
        metadata.setdefault("delivery_conditions", "")
        metadata.setdefault("payment_conditions", "")
        metadata.setdefault("logistics_cost", 0.0)
        if owner_user_id is not None:
            metadata["owner_user_id"] = int(owner_user_id)
        return metadata

    @staticmethod
    def serialize_wide_plate_lines(items: Iterable[Any]) -> list[dict[str, Any]]:
        serialized: list[dict[str, Any]] = []
        for idx, item in enumerate(items, start=1):
            if isinstance(item, (tuple, list)) and len(item) >= 2:
                serialized.append(
                    {
                        "id": f"wide-{idx}",
                        "line": str(item[0]),
                        "qty": int(item[1]),
                    }
                )
            elif isinstance(item, dict) and item.get("line"):
                serialized.append(
                    {
                        "id": str(item.get("id") or f"wide-{idx}"),
                        "line": str(item["line"]),
                        "qty": int(item.get("qty", 1) or 1),
                    }
                )
        return serialized

    @staticmethod
    def serialize_dobor_pairs(items: Iterable[Any]) -> list[dict[str, Any]]:
        serialized: list[dict[str, Any]] = []
        for idx, item in enumerate(items, start=1):
            if hasattr(item, "pair_id"):
                serialized.append(
                    {
                        "id": str(getattr(item, "pair_id", f"dobor-{idx}")),
                        "source_line": str(getattr(item, "source_line", "")),
                        "primary_line": str(getattr(item, "primary_line", "")),
                        "complement_line": str(getattr(item, "complement_line", "")),
                    }
                )
            elif isinstance(item, dict) and item.get("primary_line"):
                serialized.append(
                    {
                        "id": str(item.get("id") or f"dobor-{idx}"),
                        "source_line": str(item.get("source_line", "")),
                        "primary_line": str(item["primary_line"]),
                        "complement_line": str(item.get("complement_line", "")),
                    }
                )
        return serialized
