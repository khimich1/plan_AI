from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable

from app.core.settings import get_settings
from app.repositories.kp_repository import KpRepository
from app.repositories.manager_repository import ManagerRepository
from app.services.commercial_service import CommercialService
from app.services.draft_store import DraftStore
from app.services.execution_terms_service import ExecutionTermsService
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
        self.execution_terms_service = ExecutionTermsService()

    async def create_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        source_text, source_metadata = await self._resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )
        preview = self.commercial_service.generate_preview(text=source_text["input_text"])
        metadata = self._build_preview_metadata(
            preview=preview,
            base_metadata={},
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=source_text["input_text"],
            last_source_filename=source_text["filename"],
            plate_batches=[source_text["batch"]],
            wide_plates_resolved=not bool(preview.parse_result.wide_plate_lines),
            source_metadata=source_metadata,
        )
        draft_id = self.draft_store.save_preview(
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=preview.order_data,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

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
        draft = await self.create_draft(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )
        conditions_mode = "custom" if delivery_conditions.strip() or payment_conditions.strip() else "standard"
        return self.update_draft_meta(
            draft["draft_id"],
            manager_id=manager_id,
            client_name=client_name,
            discount_percent=discount_percent,
            conditions_mode=conditions_mode,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
        )

    async def update_draft_plates(
        self,
        draft_id: str,
        *,
        mode: str,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        source_text, source_metadata = await self._resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )
        current_text = str(metadata.get("input_text", "") or "")
        normalized_mode = mode.strip().lower()
        if normalized_mode not in {"append", "replace"}:
            raise ValueError("Некорректный режим обновления списка плит.")
        if normalized_mode == "append" and current_text:
            next_text = self._merge_plate_texts(current_text, source_text["input_text"])
            batches = list(metadata.get("plate_batches") or [])
            batches.append(source_text["batch"])
        else:
            next_text = source_text["input_text"]
            batches = [source_text["batch"]]

        preview = self.commercial_service.generate_preview(text=next_text)
        next_metadata = self._build_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=next_text,
            last_source_filename=source_text["filename"],
            plate_batches=batches,
            wide_plates_resolved=not bool(preview.parse_result.wide_plate_lines),
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        return self.get_draft_details(draft_id)

    def resolve_wide_plates(self, draft_id: str, decisions: Iterable[dict[str, Any]]) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        current_text = str(metadata.get("input_text", "") or "")
        if not current_text:
            raise ValueError("Список плит отсутствует.")

        wide_items = self._normalize_wide_plate_lines(metadata.get("wide_plate_lines", []))
        if not wide_items:
            return self.get_draft_details(draft_id)

        decisions_by_id: dict[str, dict[str, Any]] = {}
        decisions_by_line: dict[str, dict[str, Any]] = {}
        for item in decisions:
            line_id = str(item.get("line_id", "") or "").strip()
            source_line = str(item.get("source_line", "")).strip()
            if line_id:
                decisions_by_id[line_id] = item
            if not source_line:
                continue
            decisions_by_line[source_line] = item

        resolved_by_line: dict[str, dict[str, Any]] = {}
        resolved_decisions: dict[str, dict[str, Any]] = {}
        unresolved: list[str] = []
        for wide_item in wide_items:
            line_id = str(wide_item.get("id", "")).strip()
            line = str(wide_item.get("line", "")).strip()
            decision = decisions_by_id.get(line_id) if line_id else None
            if decision is None:
                decision = decisions_by_line.get(line)
            if decision is None:
                unresolved.append(line or line_id)
                continue
            resolved_decisions[line_id or line] = decision
            if line:
                resolved_by_line[line] = decision
        if unresolved:
            raise ValueError("Нужно выбрать действие для всех широких плит.")

        original_lines = [line.strip() for line in list(metadata.get("normalized_lines") or []) if line.strip()]
        if not original_lines:
            original_lines = [line.strip() for line in re.split(r"[\n;]+", current_text) if line.strip()]
        wide_set = {str(item.get("line", "")).strip() for item in wide_items if str(item.get("line", "")).strip()}
        merged_lines: list[str] = []
        for line in original_lines:
            if line not in wide_set:
                merged_lines.append(line)
                continue

            decision = resolved_by_line[line]
            action = str(decision.get("action", "")).strip().lower()
            if action == "confirm":
                merged_lines.append(line)
                continue
            if action == "exclude":
                continue
            if action == "replace":
                replacement_text = str(decision.get("replacement_text", "") or "").strip()
                replacement_lines = self._normalize_replacement_lines(replacement_text)
                if not replacement_lines:
                    raise ValueError("Для замены широкой плиты нужно указать корректный список замен.")
                merged_lines.extend(replacement_lines)
                continue
            raise ValueError("Некорректное действие для обработки широкой плиты.")

        if not merged_lines:
            raise ValueError("После обработки широких плит список стал пустым.")

        next_text = "\n".join(merged_lines)
        preview = self.commercial_service.generate_preview(text=next_text)
        next_metadata = self._build_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            plate_batches=list(metadata.get("plate_batches") or []),
            wide_plates_resolved=True,
            source_metadata={},
        )
        next_metadata["wide_plate_decisions"] = list(resolved_decisions.values())
        self.draft_store.replace_preview(
            draft_id,
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        return self.get_draft_details(draft_id)

    def update_draft_meta(
        self,
        draft_id: str,
        *,
        manager_id: int | None = None,
        client_name: str | None = None,
        discount_percent: float | None = None,
        conditions_mode: str | None = None,
        delivery_conditions: str | None = None,
        payment_conditions: str | None = None,
    ) -> dict[str, Any]:
        self._load_draft_or_raise(draft_id)
        updates: dict[str, Any] = {}
        if manager_id is not None:
            manager = self.manager_repository.get_manager(manager_id)
            if not manager:
                raise ValueError("Менеджер не найден.")
            updates.update(
                manager_id=manager["id"],
                manager_name=manager.get("fio", ""),
                manager_phone=manager.get("contact_number", ""),
                manager_email=manager.get("email", ""),
            )
        if client_name is not None:
            updates["client_name"] = client_name.strip()
        if discount_percent is not None:
            if discount_percent < 0 or discount_percent > 100:
                raise ValueError("Скидка должна быть в диапазоне от 0 до 100.")
            updates["discount_percent"] = float(discount_percent)
        if conditions_mode is not None:
            normalized_mode = conditions_mode.strip().lower()
            if normalized_mode not in {"standard", "custom"}:
                raise ValueError("Некорректный режим условий.")
            updates["conditions_mode"] = normalized_mode
        if delivery_conditions is not None:
            updates["delivery_conditions"] = delivery_conditions.strip()
        if payment_conditions is not None:
            updates["payment_conditions"] = payment_conditions.strip()
        if updates:
            updates["current_step"] = "calculate"
            self.draft_store.update_metadata(draft_id, **updates)
        return self.get_draft_details(draft_id)

    def calculate_draft(self, draft_id: str) -> dict[str, Any]:
        details = self.get_draft_details(draft_id)
        metadata = details["metadata"]
        if details["order_data"] == []:
            raise ValueError("Список плит пустой.")
        if metadata.get("wide_plate_lines") and not metadata.get("wide_plates_resolved"):
            raise ValueError("Сначала обработайте плиты шире 12 дм.")
        if not metadata.get("manager_id"):
            raise ValueError("Выберите менеджера.")
        if not str(metadata.get("client_name", "")).strip():
            raise ValueError("Укажите клиента.")
        if metadata.get("conditions_mode") == "custom":
            if not str(metadata.get("delivery_conditions", "")).strip():
                raise ValueError("Укажите условия поставки.")
            if not str(metadata.get("payment_conditions", "")).strip():
                raise ValueError("Укажите условия оплаты.")
        self.draft_store.update_metadata(draft_id, current_step="result")
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
            "saved_offer": self._normalize_saved_offer(metadata.get("saved_offer")),
            "totals": totals,
            "offer_identity": self._build_offer_identity_payload(draft_id),
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

    def save_offer(
        self,
        draft_id: str,
        *,
        execution_terms: str = "",
        status: str = "в работе",
        save_mode: str = "database",
    ) -> dict[str, Any]:
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
        saved_offer = {
            "kp_id": kp_id,
            "status": status,
            "mode": save_mode,
            "execution_terms": execution_terms,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.draft_store.update_metadata(
            draft_id,
            saved_offer=saved_offer,
            current_save_mode=save_mode,
            execution_terms=execution_terms,
        )
        totals = calculate_total_cost(
            payload["order_data"],
            float(metadata.get("discount_percent", 0.0) or 0.0),
        )
        offer_identity = self._build_offer_identity_payload(draft_id)
        return {
            "saved_offer": saved_offer,
            "totals": totals,
            "offer_identity": offer_identity,
            "result_card": self._build_result_card(
                offer_identity=offer_identity,
                metadata=metadata,
                saved_offer=saved_offer,
                totals=totals,
            ),
        }

    def save_draft(self, draft_id: str, *, mode: str, execution_terms_input: str = "") -> dict[str, Any]:
        normalized_mode = mode.strip().lower()
        if normalized_mode == "database":
            execution_terms = self.execution_terms_service.normalize(execution_terms_input)
            return self.save_offer(
                draft_id,
                execution_terms=execution_terms,
                status="в работе",
                save_mode="database",
            )
        if normalized_mode == "archive":
            return self.save_offer(
                draft_id,
                execution_terms="",
                status="в архиве",
                save_mode="archive",
            )
        if normalized_mode == "skip":
            details = self.get_draft_details(draft_id)
            skipped_offer = {
                "kp_id": None,
                "status": "не сохранено",
                "mode": "skip",
                "execution_terms": "",
                "saved_at": datetime.now().isoformat(timespec="seconds"),
            }
            self.draft_store.update_metadata(draft_id, saved_offer=skipped_offer, current_save_mode="skip")
            return {
                "saved_offer": skipped_offer,
                "totals": details["totals"],
                "offer_identity": details["offer_identity"],
                "result_card": self._build_result_card(
                    offer_identity=details["offer_identity"],
                    metadata=details["metadata"],
                    saved_offer=skipped_offer,
                    totals=details["totals"],
                ),
            }
        raise ValueError("Некорректный режим сохранения.")

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

    async def _resolve_source_input(
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
            recognized_text, source_metadata = await self._extract_text_from_image(
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

    def _build_preview_metadata(
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
                "wide_plate_lines": self._serialize_wide_plate_lines(preview.parse_result.wide_plate_lines),
                "diagnostics": list(preview.parse_result.diagnostics),
                "breakdown_tables": list(preview.breakdown_tables),
                "price_rows_count": len(preview.price_rows),
                "breakdown_tables_count": len(preview.breakdown_tables),
                "total_sum": preview.total_sum,
                "plate_batches": plate_batches,
                "wide_plates_resolved": wide_plates_resolved,
                "last_source_filename": last_source_filename,
                "current_step": "wide_plates"
                if preview.parse_result.wide_plate_lines and not wide_plates_resolved
                else metadata.get("current_step", "manager"),
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
        return metadata

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

    def _build_offer_identity_payload(self, draft_id: str) -> dict[str, str]:
        offer_number, offer_date, file_stem = self._build_offer_identity(draft_id)
        return {
            "offer_number": offer_number,
            "offer_date": offer_date,
            "file_stem": file_stem,
        }

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

    def _normalize_saved_offer(self, item: dict[str, Any] | None) -> dict[str, Any] | None:
        if not item:
            return None
        return {
            "kp_id": item.get("kp_id"),
            "status": str(item.get("status", "") or ""),
            "mode": str(item.get("mode", "database") or "database"),
            "execution_terms": str(item.get("execution_terms", "") or ""),
            "saved_at": str(item.get("saved_at", "") or ""),
        }

    def _build_result_card(
        self,
        *,
        offer_identity: dict[str, str],
        metadata: dict[str, Any],
        saved_offer: dict[str, Any],
        totals: dict[str, Any],
    ) -> dict[str, Any]:
        return {
            "kp_id": saved_offer.get("kp_id"),
            "offer_number": offer_identity["offer_number"],
            "offer_date": offer_identity["offer_date"],
            "client_name": str(metadata.get("client_name", "") or ""),
            "manager_name": str(metadata.get("manager_name", "") or ""),
            "total_amount": float(totals.get("total_with_vat", 0.0) or 0.0),
            "status": str(saved_offer.get("status", "") or ""),
            "execution_terms": str(saved_offer.get("execution_terms", "") or ""),
        }

    def _serialize_wide_plate_lines(self, items: Iterable[Any]) -> list[dict[str, Any]]:
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

    def _normalize_wide_plate_lines(self, items: Iterable[Any]) -> list[dict[str, Any]]:
        return self._serialize_wide_plate_lines(items)

    def _merge_plate_texts(self, current_text: str, next_text: str) -> str:
        parts = [current_text.strip(), next_text.strip()]
        return "\n".join(part for part in parts if part)

    def _normalize_replacement_lines(self, replacement_text: str) -> list[str]:
        if not replacement_text.strip():
            return []
        preview = self.commercial_service.generate_preview(text=replacement_text)
        return list(preview.parse_result.normalized_lines)

    def _resolve_generated_file(self, filename: str) -> Path:
        return self.settings.outputs_dir / Path(filename).name
