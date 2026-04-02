from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable

from app.core.settings import get_settings
from app.repositories.kp_repository import KpRepository
from app.repositories.manager_repository import ManagerRepository
from app.services.commercial_service import CommercialService
from app.services.draft_store import DraftStore
from app.services.file_generation_service import FileGenerationService
from core.commercial_offer_xlsx import calculate_total_cost
from core.ocr_gpt import recognize_text_smart


class CommercialWorkflowService:
    FILE_LABELS = {
        "pdf": "Коммерческое предложение (PDF)",
        "xlsx": "Коммерческое предложение (XLSX)",
        "breakdown": "Детальная разбивка (XLSX)",
        "schema": "Схема раскладки (PDF)",
    }
    DEFAULT_FILE_TYPES = ("pdf", "xlsx", "breakdown", "schema")

    def __init__(self) -> None:
        self.settings = get_settings()
        self.commercial_service = CommercialService()
        self.file_generation_service = FileGenerationService()
        self.draft_store = DraftStore()
        self.manager_repository = ManagerRepository()
        self.kp_repository = KpRepository()

    async def create_draft_from_form(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        manager_id: int,
        client_name: str,
        discount_percent: float = 0.0,
        delivery_conditions: str = "",
        payment_conditions: str = "",
    ) -> dict[str, Any]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            raise ValueError("Нужно передать текст или изображение для распознавания.")

        if image_bytes and not text_value:
            source_text, source_metadata = await self._extract_text_from_image(
                image_bytes=image_bytes,
                image_filename=image_filename,
            )
            source_type = "image"
        else:
            source_text = text_value
            source_metadata = {}
            source_type = "text"

        manager = self.manager_repository.get_manager(manager_id)
        if not manager:
            raise ValueError("Менеджер не найден.")

        preview = self.commercial_service.generate_preview(text=source_text)
        metadata = {
            "source_type": source_type,
            "original_text": text_value,
            "ocr_text": source_text if source_type == "image" else "",
            "input_text": source_text,
            "manager_id": manager["id"],
            "manager_name": manager.get("fio", ""),
            "manager_phone": manager.get("contact_number", ""),
            "manager_email": manager.get("email", ""),
            "client_name": client_name.strip(),
            "discount_percent": float(discount_percent),
            "delivery_conditions": delivery_conditions.strip(),
            "payment_conditions": payment_conditions.strip(),
            "warnings": list(preview.parse_result.warnings),
            "unparsed_lines": list(preview.parse_result.unparsed_lines),
            "normalized_text": preview.parse_result.normalized_text,
            "normalized_lines": list(preview.parse_result.normalized_lines),
            "wide_plate_lines": list(preview.parse_result.wide_plate_lines),
            "diagnostics": list(preview.parse_result.diagnostics),
            "breakdown_tables": list(preview.breakdown_tables),
            "price_rows_count": len(preview.price_rows),
            "breakdown_tables_count": len(preview.breakdown_tables),
            "total_sum": preview.total_sum,
            **source_metadata,
        }
        draft_id = self.draft_store.save_preview(
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=preview.order_data,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def get_draft_details(self, draft_id: str) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        totals = calculate_total_cost(
            payload["order_data"],
            float(metadata.get("discount_percent", 0.0) or 0.0),
        )
        public_metadata = {key: value for key, value in metadata.items() if key != "breakdown_tables"}
        return {
            "draft_id": draft_id,
            "order": payload["order"].to_dict(),
            "optimization": {
                "result": payload["optimization_context"].optimization_result,
                "total_plates": payload["optimization_context"].total_plates,
                "total_cost": payload["optimization_context"].total_cost,
            },
            "order_data": payload["order_data"],
            "metadata": public_metadata,
            "files": self._normalize_generated_files(metadata.get("generated_files", [])),
            "saved_offer": metadata.get("saved_offer"),
            "totals": totals,
        }

    def generate_files(self, draft_id: str, file_types: Iterable[str] | None = None) -> list[dict[str, str]]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        requested_types = self._normalize_file_types(file_types)
        generated_files = self._normalize_generated_files(metadata.get("generated_files", []))
        files_by_kind = {item["kind"]: item for item in generated_files}
        order_data = payload["order_data"]
        manager_name = str(metadata.get("manager_name", "") or "")
        manager_phone = str(metadata.get("manager_phone", "") or "")
        manager_email = str(metadata.get("manager_email", "") or "")
        client_name = str(metadata.get("client_name", "") or "Клиент")
        discount_percent = float(metadata.get("discount_percent", 0.0) or 0.0)
        delivery_conditions = str(metadata.get("delivery_conditions", "") or "")
        payment_conditions = str(metadata.get("payment_conditions", "") or "")
        offer_number, offer_date, file_stem = self._build_offer_identity(draft_id)

        for file_type in requested_types:
            existing = files_by_kind.get(file_type)
            if existing and self._resolve_generated_file(existing["filename"]).exists():
                continue

            if file_type == "pdf":
                output_path = self.settings.outputs_dir / f"{file_stem}.pdf"
                self.file_generation_service.generate_offer_pdf(
                    order_data=order_data,
                    output_path=str(output_path),
                    offer_number=offer_number,
                    offer_date=offer_date,
                    customer_name=client_name,
                    manager_name=manager_name,
                    manager_phone=manager_phone,
                    manager_email=manager_email,
                    discount_percent=discount_percent,
                )
                files_by_kind[file_type] = self._build_generated_file(file_type, output_path)
            elif file_type == "xlsx":
                output_path = self.settings.outputs_dir / f"{file_stem}.xlsx"
                self.file_generation_service.generate_offer_xlsx(
                    order_data=order_data,
                    output_path=str(output_path),
                    offer_number=offer_number,
                    offer_date=offer_date,
                    customer_name=client_name,
                    manager_name=manager_name,
                    manager_phone=manager_phone,
                    manager_email=manager_email,
                    discount_percent=discount_percent,
                    delivery_conditions=delivery_conditions,
                    payment_conditions=payment_conditions,
                )
                files_by_kind[file_type] = self._build_generated_file(file_type, output_path)
            elif file_type == "breakdown":
                breakdown_tables = list(metadata.get("breakdown_tables") or [])
                if not breakdown_tables:
                    continue
                output_path = self.settings.outputs_dir / f"{file_stem}_breakdown.xlsx"
                self.file_generation_service.save_breakdown(
                    breakdown_tables=breakdown_tables,
                    output_path=str(output_path),
                )
                files_by_kind[file_type] = self._build_generated_file(file_type, output_path)
            elif file_type == "schema":
                visualization_result = self.file_generation_service.generate_visualization(
                    order=payload["order"],
                    context=payload["optimization_context"],
                    output_dir=str(self.settings.outputs_dir),
                )
                if isinstance(visualization_result, tuple) and len(visualization_result) >= 2:
                    schema_path = Path(str(visualization_result[1]))
                    if schema_path.exists():
                        files_by_kind[file_type] = self._build_generated_file(file_type, schema_path)

        merged_files = [files_by_kind[key] for key in self.DEFAULT_FILE_TYPES if key in files_by_kind]
        self.draft_store.update_metadata(draft_id, generated_files=merged_files)
        return merged_files

    def save_offer(self, draft_id: str, *, execution_terms: str = "", status: str = "в работе") -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        files = self.generate_files(draft_id, ("xlsx",))
        xlsx_file = next((item for item in files if item["kind"] == "xlsx"), None)
        xlsx_path = None
        if xlsx_file:
            resolved = self._resolve_generated_file(xlsx_file["filename"])
            xlsx_path = str(resolved) if resolved.exists() else None

        kp_id = self.kp_repository.save_offer(
            creation_date=datetime.now().strftime("%d.%m.%Y"),
            customer_name=str(metadata.get("client_name", "") or "Клиент"),
            manager_name=str(metadata.get("manager_name", "") or ""),
            discount_percent=float(metadata.get("discount_percent", 0.0) or 0.0),
            delivery_conditions=str(metadata.get("delivery_conditions", "") or ""),
            payment_conditions=str(metadata.get("payment_conditions", "") or ""),
            execution_terms=execution_terms,
            status=status,
            order_data=payload["order_data"],
            xlsx_path=xlsx_path,
        )
        saved_offer = {"kp_id": kp_id, "status": status}
        self.draft_store.update_metadata(draft_id, saved_offer=saved_offer)
        totals = calculate_total_cost(
            payload["order_data"],
            float(metadata.get("discount_percent", 0.0) or 0.0),
        )
        return {"kp_id": kp_id, "status": status, "totals": totals}

    async def _extract_text_from_image(
        self,
        *,
        image_bytes: bytes,
        image_filename: str | None,
    ) -> tuple[str, dict[str, Any]]:
        suffix = Path(image_filename or "upload.jpg").suffix or ".jpg"
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
        }

    def _load_draft_or_raise(self, draft_id: str) -> dict[str, Any]:
        payload = self.draft_store.load_preview(draft_id)
        if payload is None:
            raise FileNotFoundError(f"Draft '{draft_id}' not found.")
        return payload

    def _normalize_file_types(self, file_types: Iterable[str] | None) -> list[str]:
        requested = list(file_types or self.DEFAULT_FILE_TYPES)
        normalized: list[str] = []
        for item in requested:
            key = str(item).strip().lower()
            if key in self.FILE_LABELS and key not in normalized:
                normalized.append(key)
        if not normalized:
            raise ValueError("Не выбраны типы файлов для генерации.")
        return normalized

    def _build_offer_identity(self, draft_id: str) -> tuple[str, str, str]:
        now = datetime.now()
        offer_number = f"WEB_{draft_id[:8].upper()}"
        offer_date = now.strftime("%d.%m.%Y")
        file_stem = f"kp_{draft_id[:8]}_{now.strftime('%Y%m%d_%H%M%S')}"
        return offer_number, offer_date, file_stem

    def _build_generated_file(self, kind: str, path: Path) -> dict[str, str]:
        return {
            "kind": kind,
            "filename": path.name,
            "display_name": self.FILE_LABELS[kind],
            "download_url": f"/api/v1/commercial/files/{path.name}",
        }

    def _normalize_generated_files(self, items: Iterable[dict[str, Any]]) -> list[dict[str, str]]:
        normalized: list[dict[str, str]] = []
        for item in items:
            kind = str(item.get("kind", "")).strip().lower()
            filename = Path(str(item.get("filename", "")).strip()).name
            if kind not in self.FILE_LABELS or not filename:
                continue
            normalized.append(
                {
                    "kind": kind,
                    "filename": filename,
                    "display_name": str(item.get("display_name") or self.FILE_LABELS[kind]),
                    "download_url": str(item.get("download_url") or f"/api/v1/commercial/files/{filename}"),
                }
            )
        return normalized

    def _resolve_generated_file(self, filename: str) -> Path:
        return self.settings.outputs_dir / Path(filename).name
