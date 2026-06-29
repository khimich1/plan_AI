from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from app.core.settings import get_settings
from app.services.draft_store import DraftStore
from app.services.file_generation_service import FileGenerationService
from core.commercial_offer_xlsx import DB_PATH
from core.plate_order_context import PlateOrderContext


class CommercialExportService:
    """File generation glue for commercial draft exports (PDF/XLSX/breakdown/schema)."""

    FILE_LABELS = {
        "pdf": "Коммерческое предложение (PDF)",
        "xlsx": "Коммерческое предложение (XLSX)",
        "breakdown": "Детальная разбивка (XLSX)",
        "schema": "Схема раскладки (PDF)",
    }
    DEFAULT_FILE_TYPES = ("pdf", "xlsx", "breakdown")
    ALL_FILE_TYPES = ("pdf", "xlsx", "breakdown", "schema")

    def __init__(
        self,
        *,
        draft_store: DraftStore | None = None,
        file_generation_service: FileGenerationService | None = None,
    ) -> None:
        self.settings = get_settings()
        self.draft_store = draft_store or DraftStore()
        self.file_generation_service = file_generation_service or FileGenerationService()

    def generate_files(
        self,
        draft_id: str,
        payload: dict[str, Any],
        file_types: Iterable[str] | None = None,
        *,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> list[dict[str, str]]:
        metadata = dict(payload.get("metadata", {}))
        requested_types = self.normalize_file_types(file_types)
        generated_files = self.normalize_generated_files(metadata.get("generated_files", []))
        files_by_kind = {item["kind"]: item for item in generated_files}
        schema_raw = metadata.get("schema_file")
        if schema_raw and "schema" not in files_by_kind:
            schema_items = self.normalize_generated_files([schema_raw])
            if schema_items:
                files_by_kind["schema"] = schema_items[0]
        order_data = payload["order_data"]
        from app.services.commercial_workflow_service import ensure_order_priced

        ensure_order_priced(order_data, db_path=str(DB_PATH))
        manager_name = str(metadata.get("manager_name", "") or "")
        manager_phone = str(metadata.get("manager_phone", "") or "")
        manager_email = str(metadata.get("manager_email", "") or "")
        client_name = str(metadata.get("client_name", "") or "Клиент")
        discount_percent = float(metadata.get("discount_percent", 0.0) or 0.0)
        delivery_conditions = str(metadata.get("delivery_conditions", "") or "")
        payment_conditions = str(metadata.get("payment_conditions", "") or "")
        logistics_cost = float(metadata.get("logistics_cost", 0.0) or 0.0)
        offer_number, offer_date, file_stem = self.build_offer_identity(draft_id)

        for file_type in requested_types:
            existing = files_by_kind.get(file_type)
            if (
                file_type not in {"pdf", "xlsx"}
                and existing
                and self.resolve_generated_file(existing["filename"]).exists()
            ):
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
                    logistics_cost=logistics_cost,
                    delivery_conditions=delivery_conditions or None,
                    payment_conditions=payment_conditions or None,
                )
                files_by_kind[file_type] = self.build_generated_file(draft_id, file_type, output_path)
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
                    logistics_cost=logistics_cost,
                )
                files_by_kind[file_type] = self.build_generated_file(draft_id, file_type, output_path)
            elif file_type == "breakdown":
                breakdown_tables = list(metadata.get("breakdown_tables") or [])
                if not breakdown_tables:
                    continue
                output_path = self.settings.outputs_dir / f"{file_stem}_breakdown.xlsx"
                self.file_generation_service.save_breakdown(
                    breakdown_tables=breakdown_tables,
                    output_path=str(output_path),
                )
                files_by_kind[file_type] = self.build_generated_file(draft_id, file_type, output_path)
            elif file_type == "schema":
                if plate_order_ctx is None:
                    raise ValueError(
                        "Plate order context is required for schema file generation"
                    )
                visualization_result = self.file_generation_service.generate_visualization(
                    order=payload["order"],
                    context=payload["optimization_context"],
                    ctx=plate_order_ctx,
                    output_dir=str(self.settings.outputs_dir),
                )
                if isinstance(visualization_result, tuple) and len(visualization_result) >= 2:
                    schema_path = Path(str(visualization_result[1]))
                    if schema_path.exists():
                        files_by_kind[file_type] = self.build_generated_file(draft_id, file_type, schema_path)

        merged_files = [files_by_kind[key] for key in self.DEFAULT_FILE_TYPES if key in files_by_kind]
        update_payload: dict[str, Any] = {"generated_files": merged_files}
        if "schema" in files_by_kind:
            update_payload["schema_file"] = files_by_kind["schema"]
        self.draft_store.update_metadata(draft_id, **update_payload)
        return [files_by_kind[key] for key in requested_types if key in files_by_kind]

    def collect_draft_files(self, metadata: dict[str, Any], draft_id: str) -> list[dict[str, str]]:
        files = self.normalize_generated_files(metadata.get("generated_files", []))
        schema_raw = metadata.get("schema_file")
        if schema_raw:
            schema_items = self.normalize_generated_files([schema_raw])
            if schema_items and not any(item["kind"] == "schema" for item in files):
                schema_item = schema_items[0]
                filename = schema_item["filename"]
                files.append(
                    self.build_generated_file(draft_id, "schema", Path(filename))
                    if "draft_id=" not in schema_item.get("download_url", "")
                    else schema_item
                )
        return files

    def get_or_generate_file(self, safe_filename: str) -> Path:
        """Path under configured ``outputs_dir`` for a generated file basename (no subpaths)."""
        return self.resolve_generated_file(safe_filename)

    def normalize_file_types(self, file_types: Iterable[str] | None) -> list[str]:
        requested = list(file_types or self.DEFAULT_FILE_TYPES)
        normalized: list[str] = []
        for item in requested:
            key = str(item).strip().lower()
            if key in self.FILE_LABELS and key not in normalized:
                normalized.append(key)
        if not normalized:
            raise ValueError("Не выбраны типы файлов для генерации.")
        return normalized

    def build_offer_identity(self, draft_id: str) -> tuple[str, str, str]:
        now = datetime.now()
        offer_number = f"WEB_{draft_id[:8].upper()}"
        offer_date = now.strftime("%d.%m.%Y")
        file_stem = f"kp_{draft_id[:8]}_{now.strftime('%Y%m%d_%H%M%S')}"
        return offer_number, offer_date, file_stem

    def build_offer_identity_payload(self, draft_id: str) -> dict[str, str]:
        offer_number, offer_date, file_stem = self.build_offer_identity(draft_id)
        return {
            "offer_number": offer_number,
            "offer_date": offer_date,
            "file_stem": file_stem,
        }

    def build_generated_file(self, draft_id: str, kind: str, path: Path) -> dict[str, str]:
        name = path.name
        return {
            "kind": kind,
            "filename": name,
            "display_name": self.FILE_LABELS[kind],
            "download_url": f"/api/v1/commercial/files/{name}?draft_id={draft_id}",
        }

    def normalize_generated_files(self, items: Iterable[dict[str, Any]]) -> list[dict[str, str]]:
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
                    "download_url": str(
                        item.get("download_url") or f"/api/v1/commercial/files/{filename}"
                    ),
                }
            )
        return normalized

    def resolve_generated_file(self, filename: str) -> Path:
        return self.settings.outputs_dir / Path(filename).name
