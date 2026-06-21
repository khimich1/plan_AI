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


class CommercialDraftService:
    """Draft ingest helpers: OCR/text source resolution and preview metadata assembly."""

    def __init__(self, *, commercial_service: CommercialService | None = None) -> None:
        self.settings = get_settings()
        self.commercial_service = commercial_service or CommercialService()

    async def extract_text_from_image(
        self,
        *,
        image_bytes: bytes,
        image_filename: str | None,
    ) -> tuple[str, dict[str, Any]]:
        suffix = safe_ocr_temp_suffix(image_filename)
        with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(image_bytes)
            tmp_path = Path(tmp_file.name)

        recognition_mode = (self.settings.ocr_recognition_mode or "full_gpt").strip().lower()
        if recognition_mode not in {"full_gpt", "hybrid"}:
            recognition_mode = "full_gpt"

        try:
            result = await recognize_text_smart(
                str(tmp_path),
                force_gpt=(recognition_mode == "full_gpt"),
                show_cost=True,
                mode=recognition_mode,
                verify_enabled=self.settings.ocr_verify_enabled,
            )
        finally:
            tmp_path.unlink(missing_ok=True)

        recognized_text = str((result or {}).get("text", "")).strip()
        if not recognized_text:
            raise ValueError("Не удалось распознать текст на изображении.")
        return recognized_text, {
            "ocr_recognition_mode": recognition_mode,
            "ocr_cost_usd": float((result or {}).get("cost_usd", 0.0) or 0.0),
            "ocr_plates": list((result or {}).get("plates") or []),
            "ocr_draft_plates": list((result or {}).get("draft_plates") or []),
            "ocr_corrections": list((result or {}).get("corrections") or []),
            "ocr_verify_applied": bool((result or {}).get("verify_applied")),
            "ocr_verify_failed": bool((result or {}).get("verify_failed")),
            "ocr_method": str((result or {}).get("method") or "GPT-4o"),
            "ocr_row_count_on_image": (result or {}).get("row_count_on_image"),
        }

    async def resolve_source_input(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            raise ValueError("Нужно передать текст или изображение для распознавания.")
        if image_bytes and not text_value:
            recognized_text, source_metadata = await self.extract_text_from_image(
                image_bytes=image_bytes,
                image_filename=image_filename,
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
        metadata = dict(base_metadata)
        metadata.update(
            {
                "source_type": source_type,
                "original_text": original_text,
                "ocr_text": ocr_text,
                "input_text": input_text,
                "accumulated_text": input_text,
                "warnings": list(preview.parse_result.warnings),
                "unparsed_lines": list(preview.parse_result.unparsed_lines),
                "normalized_text": preview.parse_result.normalized_text,
                "normalized_lines": list(preview.parse_result.normalized_lines),
                "wide_plate_lines": self.serialize_wide_plate_lines(preview.parse_result.wide_plate_lines),
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
                **source_metadata,
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
