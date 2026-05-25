from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable

from app.core.settings import get_settings
from app.schemas.commercial import WizardNextRequiredAction, WizardStepId
from app.repositories.kp_repository import KpRepository
from app.repositories.manager_repository import ManagerRepository
from app.services.commercial_calculation_service import CommercialCalculationService
from app.services.commercial_service import CommercialService
from app.services.draft_store import DraftStore, UnsafeDraftIdError
from app.services.execution_terms_service import ExecutionTermsService
from app.services.file_generation_service import FileGenerationService
from core.ocr_gpt import recognize_text_smart


_ALLOWED_OCR_IMAGE_SUFFIXES = frozenset({".jpg", ".jpeg", ".png", ".webp", ".gif", ".pdf"})


def _safe_ocr_temp_suffix(image_filename: str | None) -> str:
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


class CommercialWorkflowService:
    FILE_LABELS = {
        "pdf": "Коммерческое предложение (PDF)",
        "xlsx": "Коммерческое предложение (XLSX)",
        "breakdown": "Детальная разбивка (XLSX)",
        "schema": "Схема раскладки (PDF)",
    }
    # Схема раскладки (matplotlib) — отдельно: при включении в file_types на слабом VPS возможен OOM.
    DEFAULT_FILE_TYPES = ("pdf", "xlsx", "breakdown")
    ALL_FILE_TYPES = ("pdf", "xlsx", "breakdown", "schema")

    def __init__(self) -> None:
        self.settings = get_settings()
        self.commercial_service = CommercialService()
        self.file_generation_service = FileGenerationService()
        self.draft_store = DraftStore()
        self.manager_repository = ManagerRepository()
        self.kp_repository = KpRepository()
        self.execution_terms_service = ExecutionTermsService()
        self.calculation_service = CommercialCalculationService()

    def _wide_lines_blocking(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.wide_lines_blocking(metadata)

    def _meta_ready_for_calculate(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.meta_ready_for_calculate(metadata)

    def _normalize_stored_step(self, metadata: dict[str, Any]) -> WizardStepId:
        raw = str(metadata.get("current_step") or "").strip().lower()
        aliases = {"wide_plates": WizardStepId.wide_plates.value, "calculate": WizardStepId.client.value}
        raw = aliases.get(raw, raw)
        try:
            return WizardStepId(raw) if raw else WizardStepId.plates
        except ValueError:
            return WizardStepId.plates

    def _wizard_step_after_plate_snapshot(self, metadata: dict[str, Any], order_data: list[Any]) -> WizardStepId:
        return WizardStepId.plates

    def _persist_wizard_step(self, draft_id: str, step: WizardStepId) -> None:
        self.draft_store.update_metadata(draft_id, current_step=step.value)

    def infer_wizard_current_step(self, payload: dict[str, Any]) -> WizardStepId:
        """Эффективный шаг для UI: хранимый шаг + защита от «перепрыгивания» узких плит + валидность result."""
        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []
        stored = self._normalize_stored_step(metadata)

        if stored == WizardStepId.result:
            if (
                order_data
                and not self._wide_lines_blocking(metadata)
                and self._meta_ready_for_calculate(metadata)
            ):
                return WizardStepId.result
            return WizardStepId.client

        if order_data and self._wide_lines_blocking(metadata):
            if stored in (WizardStepId.manager, WizardStepId.client, WizardStepId.result):
                return WizardStepId.wide_plates
            if stored == WizardStepId.wide_plates:
                return WizardStepId.wide_plates

        return stored

    def _infer_next_required_action(
        self,
        payload: dict[str, Any],
        effective_step: WizardStepId,
    ) -> WizardNextRequiredAction:
        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []

        if not order_data:
            return WizardNextRequiredAction.ingest_plates

        if self._wide_lines_blocking(metadata):
            return WizardNextRequiredAction.resolve_wide_plates

        if not metadata.get("manager_id"):
            return WizardNextRequiredAction.select_manager

        if not self._meta_ready_for_calculate(metadata):
            return WizardNextRequiredAction.complete_client_terms

        if effective_step == WizardStepId.result:
            return WizardNextRequiredAction.none

        return WizardNextRequiredAction.post_calculate

    def _infer_can_proceed_to(
        self,
        payload: dict[str, Any],
        effective_step: WizardStepId,
        next_action: WizardNextRequiredAction,
    ) -> list[WizardStepId]:
        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []

        if effective_step == WizardStepId.plates:
            if not order_data:
                return []
            if self._wide_lines_blocking(metadata):
                return [WizardStepId.wide_plates]
            return [WizardStepId.manager]

        if effective_step == WizardStepId.wide_plates:
            return []

        if effective_step == WizardStepId.manager:
            if metadata.get("manager_id"):
                return [WizardStepId.client]
            return []

        if effective_step == WizardStepId.client:
            if next_action == WizardNextRequiredAction.post_calculate:
                return []
            return []

        return []

    def _collect_wizard_validation_errors(
        self,
        payload: dict[str, Any],
        next_action: WizardNextRequiredAction,
    ) -> list[str]:
        """Сообщения, согласованные с проверками ``calculate_draft`` / ``next_required_action``."""
        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []

        if next_action == WizardNextRequiredAction.ingest_plates:
            return ["Список плит пустой."]
        if next_action == WizardNextRequiredAction.resolve_wide_plates:
            return ["Сначала обработайте плиты шире 12 дм."]
        if next_action == WizardNextRequiredAction.select_manager:
            return ["Выберите менеджера."]
        if next_action == WizardNextRequiredAction.complete_client_terms:
            errors: list[str] = []
            if not order_data:
                errors.append("Список плит пустой.")
            if metadata.get("wide_plate_lines") and not metadata.get("wide_plates_resolved"):
                errors.append("Сначала обработайте плиты шире 12 дм.")
            if not metadata.get("manager_id"):
                errors.append("Выберите менеджера.")
            if not str(metadata.get("client_name", "")).strip():
                errors.append("Укажите клиента.")
            if metadata.get("conditions_mode") == "custom":
                if not str(metadata.get("delivery_conditions", "")).strip():
                    errors.append("Укажите условия поставки.")
                if not str(metadata.get("payment_conditions", "")).strip():
                    errors.append("Укажите условия оплаты.")
            return errors
        return []

    def build_wizard_state(self, payload: dict[str, Any]) -> dict[str, Any]:
        current = self.infer_wizard_current_step(payload)
        next_action = self._infer_next_required_action(payload, current)
        can_proceed = self._infer_can_proceed_to(payload, current, next_action)
        validation_errors = self._collect_wizard_validation_errors(payload, next_action)
        return {
            "current_step": current,
            "can_proceed_to": can_proceed,
            "next_required_action": next_action,
            "validation_errors": validation_errors,
        }

    async def create_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
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
            owner_user_id=owner_user_id,
        )
        draft_id = self.draft_store.save_preview(
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=preview.order_data,
            metadata=metadata,
        )
        payload_snap = self._load_draft_or_raise(draft_id)
        plates_step = self._wizard_step_after_plate_snapshot(
            dict(payload_snap.get("metadata", {})),
            payload_snap["order_data"],
        )
        self._persist_wizard_step(draft_id, plates_step)
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
        owner_user_id: int,
    ) -> dict[str, Any]:
        draft = await self.create_draft(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            owner_user_id=owner_user_id,
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
        payload_snap = self._load_draft_or_raise(draft_id)
        plates_step = self._wizard_step_after_plate_snapshot(
            dict(payload_snap.get("metadata", {})),
            payload_snap["order_data"],
        )
        self._persist_wizard_step(draft_id, plates_step)
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
        self._persist_wizard_step(draft_id, WizardStepId.manager)
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
        logistics_cost: float | None = None,
    ) -> dict[str, Any]:
        payload_before = self._load_draft_or_raise(draft_id)
        prev_step = self._normalize_stored_step(dict(payload_before.get("metadata") or {}))

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
        if logistics_cost is not None:
            if logistics_cost < 0:
                raise ValueError("Стоимость рейса не может быть отрицательной.")
            updates["logistics_cost"] = float(logistics_cost)
        if updates:
            self.draft_store.update_metadata(draft_id, **updates)

        payload_after = self._load_draft_or_raise(draft_id)
        md = dict(payload_after.get("metadata") or {})

        financial_keys = {"discount_percent", "logistics_cost"}
        if updates:
            if prev_step == WizardStepId.result and set(updates.keys()).issubset(financial_keys):
                self._persist_wizard_step(draft_id, WizardStepId.result)
            elif "manager_id" in updates and not (
                {"client_name", "conditions_mode", "delivery_conditions", "payment_conditions"} & set(updates.keys())
            ):
                self._persist_wizard_step(draft_id, WizardStepId.manager)
            elif self._meta_ready_for_calculate(md):
                self._persist_wizard_step(draft_id, WizardStepId.client)
            elif md.get("manager_id"):
                self._persist_wizard_step(draft_id, WizardStepId.client)
            else:
                self._persist_wizard_step(draft_id, WizardStepId.manager)

        return self.get_draft_details(draft_id)

    def calculate_draft(self, draft_id: str) -> dict[str, Any]:
        details = self.get_draft_details(draft_id)
        metadata = details["metadata"]
        self.calculation_service.validate_calculate_prerequisites(
            order_data=list(details["order_data"]),
            metadata=dict(metadata),
        )
        self.draft_store.update_metadata(draft_id, current_step=WizardStepId.result.value)
        return self.get_draft_details(draft_id)

    def get_draft_details(self, draft_id: str) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        totals = self.calculation_service.compute_totals(
            payload["order_data"],
            discount_percent=float(metadata.get("discount_percent", 0.0) or 0.0),
            logistics_cost=float(metadata.get("logistics_cost", 0.0) or 0.0),
        )
        public_metadata = {
            key: value
            for key, value in metadata.items()
            if key not in ("breakdown_tables", "owner_user_id", "predicted_kp_id")
        }
        wizard_state = self.build_wizard_state(payload)
        public_metadata["current_step"] = wizard_state["current_step"].value

        return {
            "draft_id": draft_id,
            "order": payload["order"].to_dict(),
            "optimization": {
                "result": payload["optimization_context"].optimization_result,
                "total_plates": payload["optimization_context"].total_plates,
                "total_cost": payload["optimization_context"].total_cost,
                "status": payload["optimization_context"].optimization_status,
                "success": payload["optimization_context"].optimization_success,
                "error_code": payload["optimization_context"].optimization_error_code,
                "error_message": payload["optimization_context"].optimization_error_message,
            },
            "order_data": payload["order_data"],
            "metadata": public_metadata,
            "wizard_state": wizard_state,
            "files": self._collect_draft_files(metadata, draft_id),
            "saved_offer": self._normalize_saved_offer(metadata.get("saved_offer")),
            "totals": totals,
            "offer_identity": self._build_offer_identity_payload(draft_id, metadata),
        }

    def generate_files(self, draft_id: str, file_types: Iterable[str] | None = None) -> list[dict[str, str]]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        requested_types = self._normalize_file_types(file_types)
        generated_files = self._normalize_generated_files(metadata.get("generated_files", []))
        files_by_kind = {item["kind"]: item for item in generated_files}
        schema_raw = metadata.get("schema_file")
        if schema_raw and "schema" not in files_by_kind:
            schema_items = self._normalize_generated_files([schema_raw])
            if schema_items:
                files_by_kind["schema"] = schema_items[0]
        order_data = payload["order_data"]
        manager_name = str(metadata.get("manager_name", "") or "")
        manager_phone = str(metadata.get("manager_phone", "") or "")
        manager_email = str(metadata.get("manager_email", "") or "")
        client_name = str(metadata.get("client_name", "") or "Клиент")
        discount_percent = float(metadata.get("discount_percent", 0.0) or 0.0)
        delivery_conditions = str(metadata.get("delivery_conditions", "") or "")
        payment_conditions = str(metadata.get("payment_conditions", "") or "")
        logistics_cost = float(metadata.get("logistics_cost", 0.0) or 0.0)
        offer_number, offer_date, file_stem, kp_db_id = self._build_offer_identity(
            draft_id,
            metadata,
            persist_predicted_kp_id=True,
        )

        for file_type in requested_types:
            existing = files_by_kind.get(file_type)
            if (
                file_type not in {"pdf", "xlsx"}
                and existing
                and self._resolve_generated_file(existing["filename"]).exists()
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
                    kp_db_id=kp_db_id,
                )
                files_by_kind[file_type] = self._build_generated_file(draft_id, file_type, output_path)
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
                    kp_db_id=kp_db_id,
                )
                files_by_kind[file_type] = self._build_generated_file(draft_id, file_type, output_path)
            elif file_type == "breakdown":
                breakdown_tables = list(metadata.get("breakdown_tables") or [])
                if not breakdown_tables:
                    continue
                output_path = self.settings.outputs_dir / f"{file_stem}_breakdown.xlsx"
                self.file_generation_service.save_breakdown(
                    breakdown_tables=breakdown_tables,
                    output_path=str(output_path),
                )
                files_by_kind[file_type] = self._build_generated_file(draft_id, file_type, output_path)
            elif file_type == "schema":
                visualization_result = self.file_generation_service.generate_visualization(
                    order=payload["order"],
                    context=payload["optimization_context"],
                    output_dir=str(self.settings.outputs_dir),
                )
                if isinstance(visualization_result, tuple) and len(visualization_result) >= 2:
                    schema_path = Path(str(visualization_result[1]))
                    if schema_path.exists():
                        files_by_kind[file_type] = self._build_generated_file(draft_id, file_type, schema_path)

        merged_files = [files_by_kind[key] for key in self.DEFAULT_FILE_TYPES if key in files_by_kind]
        update_payload: dict[str, Any] = {"generated_files": merged_files}
        if "schema" in files_by_kind:
            update_payload["schema_file"] = files_by_kind["schema"]
        self.draft_store.update_metadata(draft_id, **update_payload)
        return [files_by_kind[key] for key in requested_types if key in files_by_kind]

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
            logistics_cost=float(metadata.get("logistics_cost", 0.0) or 0.0),
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
        totals = self.calculation_service.compute_totals(
            payload["order_data"],
            discount_percent=float(metadata.get("discount_percent", 0.0) or 0.0),
            logistics_cost=float(metadata.get("logistics_cost", 0.0) or 0.0),
        )
        metadata["saved_offer"] = saved_offer
        offer_identity = self._build_offer_identity_payload(draft_id, metadata)
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
            raw = (execution_terms_input or "").strip()
            execution_terms = self.execution_terms_service.normalize(raw) if raw else ""
            return self.save_offer(
                draft_id,
                execution_terms=execution_terms,
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
        suffix = _safe_ocr_temp_suffix(image_filename)
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
        try:
            payload = self.draft_store.load_preview(draft_id)
        except UnsafeDraftIdError:
            raise FileNotFoundError(f"Draft '{draft_id}' not found.") from None
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
                "wide_plate_lines": self._serialize_wide_plate_lines(preview.parse_result.wide_plate_lines),
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

    def _resolve_kp_number_for_draft(
        self,
        metadata: dict[str, Any],
        draft_id: str,
        *,
        persist_predicted_kp_id: bool,
    ) -> int:
        saved = metadata.get("saved_offer") or {}
        saved_id = saved.get("kp_id")
        if saved_id is not None:
            return int(saved_id)

        predicted = metadata.get("predicted_kp_id")
        if predicted is not None:
            return int(predicted)

        next_id = self.kp_repository.get_next_kp_number()
        if persist_predicted_kp_id:
            self.draft_store.update_metadata(draft_id, predicted_kp_id=next_id)
            metadata["predicted_kp_id"] = next_id
        return next_id

    def _build_offer_identity(
        self,
        draft_id: str,
        metadata: dict[str, Any] | None = None,
        *,
        persist_predicted_kp_id: bool = False,
    ) -> tuple[str, str, str, int]:
        if metadata is None:
            payload = self._load_draft_or_raise(draft_id)
            meta: dict[str, Any] = dict(payload.get("metadata") or {})
        else:
            meta = metadata
        kp_id = self._resolve_kp_number_for_draft(
            meta,
            draft_id,
            persist_predicted_kp_id=persist_predicted_kp_id,
        )
        now = datetime.now()
        offer_number = str(kp_id)
        offer_date = now.strftime("%d.%m.%Y")
        file_stem = f"kp_{kp_id}_{now.strftime('%Y%m%d_%H%M%S')}"
        return offer_number, offer_date, file_stem, kp_id

    def _build_offer_identity_payload(
        self,
        draft_id: str,
        metadata: dict[str, Any] | None = None,
    ) -> dict[str, str]:
        offer_number, offer_date, file_stem, _ = self._build_offer_identity(
            draft_id,
            metadata,
            persist_predicted_kp_id=False,
        )
        return {
            "offer_number": offer_number,
            "offer_date": offer_date,
            "file_stem": file_stem,
        }

    def _build_generated_file(self, draft_id: str, kind: str, path: Path) -> dict[str, str]:
        name = path.name
        return {
            "kind": kind,
            "filename": name,
            "display_name": self.FILE_LABELS[kind],
            "download_url": f"/api/v1/commercial/files/{name}?draft_id={draft_id}",
        }

    def _collect_draft_files(self, metadata: dict[str, Any], draft_id: str) -> list[dict[str, str]]:
        files = self._normalize_generated_files(metadata.get("generated_files", []))
        schema_raw = metadata.get("schema_file")
        if schema_raw:
            schema_items = self._normalize_generated_files([schema_raw])
            if schema_items and not any(item["kind"] == "schema" for item in files):
                schema_item = schema_items[0]
                filename = schema_item["filename"]
                files.append(
                    self._build_generated_file(draft_id, "schema", Path(filename))
                    if "draft_id=" not in schema_item.get("download_url", "")
                    else schema_item
                )
        return files

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
                    "download_url": str(
                        item.get("download_url") or f"/api/v1/commercial/files/{filename}"
                    ),
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

    def get_or_generate_file(self, safe_filename: str) -> Path:
        """Path under configured ``outputs_dir`` for a generated file basename (no subpaths)."""
        return self._resolve_generated_file(safe_filename)

    def _resolve_generated_file(self, filename: str) -> Path:
        return self.settings.outputs_dir / Path(filename).name
