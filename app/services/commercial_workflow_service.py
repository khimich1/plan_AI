from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable

from app.core.settings import get_settings
from app.schemas.commercial import WizardStepId
from app.repositories.kp_repository import KpRepository
from app.repositories.manager_repository import ManagerRepository
from app.services.commercial_calculation_service import CommercialCalculationService
from app.services.commercial_draft_service import CommercialDraftService, _safe_ocr_temp_suffix
from app.services.commercial_export_service import CommercialExportService
from app.services.commercial_bridge_pile_service import CommercialBridgePileService
from app.services.commercial_fbs_service import CommercialFbsService
from app.services.commercial_march_service import CommercialMarchService
from app.services.commercial_pile_service import CommercialPileService
from app.services.commercial_step_service import CommercialStepService
from app.services.commercial_service import CommercialService
from app.services.commercial_wizard_step_service import CommercialWizardStepService
from app.services.draft_store import DraftStore, UnsafeDraftIdError
from app.services.execution_terms_service import ExecutionTermsService
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from core.commercial_offer_xlsx import DB_PATH
from core.commercial_pricing import ensure_order_priced
from core.bridge_pile_price_db import list_available_grades
from core.fbs_price_db import list_available_grades as list_fbs_available_grades
from core.ocr_gpt import (
    apply_bridge_piles_with_ai,
    apply_fbs_with_ai,
    apply_marches_with_ai,
    apply_piles_with_ai,
    apply_plates_with_ai,
    apply_steps_with_ai,
)
from core.plate_order_context import PlateOrderContext


class CommercialWorkflowService:
    FILE_LABELS = CommercialExportService.FILE_LABELS
    DEFAULT_FILE_TYPES = CommercialExportService.DEFAULT_FILE_TYPES
    ALL_FILE_TYPES = CommercialExportService.ALL_FILE_TYPES

    def __init__(self) -> None:
        self.settings = get_settings()
        self.commercial_service = CommercialService()
        self.draft_store = DraftStore()
        self.manager_repository = ManagerRepository()
        self.kp_repository = KpRepository()
        self.execution_terms_service = ExecutionTermsService()
        self.calculation_service = CommercialCalculationService()
        self.draft_service = CommercialDraftService(commercial_service=self.commercial_service)
        self.pile_service = CommercialPileService()
        self.march_service = CommercialMarchService()
        self.bridge_pile_service = CommercialBridgePileService()
        self.fbs_service = CommercialFbsService()
        self.step_service_product = CommercialStepService()
        self.export_service = CommercialExportService(draft_store=self.draft_store)
        self.file_generation_service = self.export_service.file_generation_service
        self.step_service = CommercialWizardStepService(
            calculation_service=self.calculation_service,
            draft_store=self.draft_store,
        )

    def _wide_lines_blocking(self, metadata: dict[str, Any]) -> bool:
        return self.step_service.wide_lines_blocking(metadata)

    def _meta_ready_for_calculate(self, metadata: dict[str, Any]) -> bool:
        return self.step_service.meta_ready_for_calculate(metadata)

    def _normalize_stored_step(self, metadata: dict[str, Any]) -> WizardStepId:
        return self.step_service.normalize_stored_step(metadata)

    def _wizard_step_after_plate_snapshot(self, metadata: dict[str, Any], order_data: list[Any]) -> WizardStepId:
        return self.step_service.wizard_step_after_plate_snapshot(metadata, order_data)

    def _persist_wizard_step(self, draft_id: str, step: WizardStepId) -> None:
        self.step_service.persist_wizard_step(draft_id, step)

    def infer_wizard_current_step(self, payload: dict[str, Any]) -> WizardStepId:
        return self.step_service.infer_wizard_current_step(payload)

    def build_wizard_state(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self.step_service.build_wizard_state(payload)

    async def create_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
        plate_order_ctx: PlateOrderContext,
        product_type: str = "plates",
    ) -> dict[str, Any]:
        normalized_product_type = (product_type or "plates").strip().lower()
        if normalized_product_type not in {"plates", "piles", "steps", "marches", "bridge_piles", "fbs"}:
            raise ValueError("Некорректный тип продукта.")
        if normalized_product_type == "piles":
            return await self._create_pile_draft(
                text=text,
                image_bytes=image_bytes,
                image_filename=image_filename,
                owner_user_id=owner_user_id,
            )
        if normalized_product_type == "bridge_piles":
            return await self._create_bridge_pile_draft(
                text=text,
                image_bytes=image_bytes,
                image_filename=image_filename,
                owner_user_id=owner_user_id,
            )
        if normalized_product_type == "fbs":
            return await self._create_fbs_draft(
                text=text,
                image_bytes=image_bytes,
                image_filename=image_filename,
                owner_user_id=owner_user_id,
            )
        if normalized_product_type == "marches":
            return await self._create_march_draft(
                text=text,
                image_bytes=image_bytes,
                image_filename=image_filename,
                owner_user_id=owner_user_id,
            )
        if normalized_product_type == "steps":
            return await self._create_step_draft(
                text=text,
                image_bytes=image_bytes,
                image_filename=image_filename,
                owner_user_id=owner_user_id,
            )

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )
        preview = self.commercial_service.generate_preview(
            text=source_text["input_text"],
            plate_order_ctx=plate_order_ctx,
        )
        metadata = self.draft_service.build_preview_metadata(
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

    async def _create_pile_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
    ) -> dict[str, Any]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            metadata = {
                "product_type": "piles",
                "owner_user_id": owner_user_id,
                "current_step": WizardStepId.piles.value,
                "wide_plates_resolved": True,
                "default_concrete_grade": "B25",
            }
            draft_id = self.draft_store.save_preview(
                order=PlateOrder(),
                optimization_context=OptimizationContext(order=PlateOrder()),
                order_data=[],
                metadata=metadata,
            )
            return self.get_draft_details(draft_id)

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="piles",
        )
        preview = self.pile_service.generate_preview(
            source_text["input_text"],
            db_path=str(DB_PATH),
        )
        batches = [source_text["batch"]]
        metadata = self.draft_service.build_pile_preview_metadata(
            preview=preview,
            base_metadata={"product_type": "piles"},
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=source_text["input_text"],
            last_source_filename=source_text["filename"],
            pile_batches=batches,
            source_metadata=source_metadata,
            owner_user_id=owner_user_id,
        )
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.piles)
        return self.get_draft_details(draft_id)

    async def _create_march_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
    ) -> dict[str, Any]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            metadata = {
                "product_type": "marches",
                "owner_user_id": owner_user_id,
                "current_step": WizardStepId.marches.value,
                "wide_plates_resolved": True,
                "default_concrete_grade": "B25",
            }
            draft_id = self.draft_store.save_preview(
                order=PlateOrder(),
                optimization_context=OptimizationContext(order=PlateOrder()),
                order_data=[],
                metadata=metadata,
            )
            return self.get_draft_details(draft_id)

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="marches",
        )
        preview = self.march_service.generate_preview(
            source_text["input_text"],
            db_path=str(DB_PATH),
        )
        batches = [source_text["batch"]]
        metadata = self.draft_service.build_march_preview_metadata(
            preview=preview,
            base_metadata={"product_type": "marches"},
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=source_text["input_text"],
            last_source_filename=source_text["filename"],
            march_batches=batches,
            source_metadata=source_metadata,
            owner_user_id=owner_user_id,
        )
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.marches)
        return self.get_draft_details(draft_id)

    async def _create_bridge_pile_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
    ) -> dict[str, Any]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            metadata = {
                "product_type": "bridge_piles",
                "owner_user_id": owner_user_id,
                "current_step": WizardStepId.bridge_piles.value,
                "wide_plates_resolved": True,
                "default_concrete_grade": "B25",
            }
            draft_id = self.draft_store.save_preview(
                order=PlateOrder(),
                optimization_context=OptimizationContext(order=PlateOrder()),
                order_data=[],
                metadata=metadata,
            )
            return self.get_draft_details(draft_id)

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="bridge_piles",
        )
        preview = self.bridge_pile_service.generate_preview(
            source_text["input_text"],
            db_path=str(DB_PATH),
        )
        batches = [source_text["batch"]]
        metadata = self.draft_service.build_bridge_pile_preview_metadata(
            preview=preview,
            base_metadata={"product_type": "bridge_piles"},
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=source_text["input_text"],
            last_source_filename=source_text["filename"],
            bridge_pile_batches=batches,
            source_metadata=source_metadata,
            owner_user_id=owner_user_id,
        )
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.bridge_piles)
        return self.get_draft_details(draft_id)

    async def _create_step_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
    ) -> dict[str, Any]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            metadata = {
                "product_type": "steps",
                "owner_user_id": owner_user_id,
                "current_step": WizardStepId.steps.value,
                "wide_plates_resolved": True,
            }
            draft_id = self.draft_store.save_preview(
                order=PlateOrder(),
                optimization_context=OptimizationContext(order=PlateOrder()),
                order_data=[],
                metadata=metadata,
            )
            return self.get_draft_details(draft_id)

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="steps",
        )
        preview = self.step_service_product.generate_preview(
            source_text["input_text"],
            db_path=str(DB_PATH),
        )
        batches = [source_text["batch"]]
        metadata = self.draft_service.build_step_preview_metadata(
            preview=preview,
            base_metadata={"product_type": "steps"},
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=source_text["input_text"],
            last_source_filename=source_text["filename"],
            step_batches=batches,
            source_metadata=source_metadata,
            owner_user_id=owner_user_id,
        )
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.steps)
        return self.get_draft_details(draft_id)

    async def update_draft_piles(
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
        if str(metadata.get("product_type", "plates")).lower() != "piles":
            raise ValueError("Черновик не является КП на сваи.")

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="piles",
        )
        current_text = str(metadata.get("input_text", "") or "")
        normalized_mode = mode.strip().lower()
        if normalized_mode not in {"append", "replace"}:
            raise ValueError("Некорректный режим обновления списка свай.")
        if normalized_mode == "append" and current_text:
            next_text = self._merge_plate_texts(current_text, source_text["input_text"])
            batches = list(metadata.get("pile_batches") or [])
            batches.append(source_text["batch"])
        else:
            next_text = source_text["input_text"]
            batches = [source_text["batch"]]

        preview = self.pile_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_pile_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=next_text,
            last_source_filename=source_text["filename"],
            pile_batches=batches,
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.piles)
        return self.get_draft_details(draft_id)

    async def apply_ai_piles_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "piles":
            raise ValueError("ИИ-редактирование свай доступно только для КП на сваи.")

        instruction_value = (instruction or "").strip()
        if len(instruction_value) < 3:
            raise ValueError("Инструкция для ИИ должна содержать минимум 3 символа.")

        current_text = str(metadata.get("input_text", "") or "")
        tmp_path: Path | None = None
        try:
            if image_bytes:
                suffix = _safe_ocr_temp_suffix(image_filename)
                with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(image_bytes)
                    tmp_path = Path(tmp_file.name)
                result = await apply_piles_with_ai(
                    current_piles_text=current_text,
                    user_instruction=instruction_value,
                    image_path=str(tmp_path),
                )
            else:
                result = await apply_piles_with_ai(
                    current_piles_text=current_text,
                    user_instruction=instruction_value,
                )
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

        recognized_text = str((result or {}).get("text", "")).strip()
        if not recognized_text:
            raise ValueError("ИИ не смог обработать список свай. Попробуйте уточнить инструкцию.")

        source_metadata = {
            "ai_applied": True,
            "last_ai_instruction": instruction_value,
            "ai_cost_usd": float((result or {}).get("cost_usd", 0.0) or 0.0),
            "ai_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ai_plates": list((result or {}).get("plates") or []),
            "ocr_plates": list((result or {}).get("plates") or []),
            "ocr_draft_plates": list((result or {}).get("draft_plates") or []),
            "ocr_corrections": list((result or {}).get("corrections") or []),
            "ocr_verify_applied": False,
            "ocr_verify_failed": False,
            "ocr_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ocr_row_count_on_image": (result or {}).get("row_count_on_image"),
        }
        ai_batch = {
            "source_type": "ai",
            "original_text": instruction_value,
            "normalized_text": recognized_text,
            "ocr_text": recognized_text,
            "filename": image_filename or "",
        }
        preview = self.pile_service.generate_preview(
            recognized_text,
            db_path=str(DB_PATH),
        )
        next_metadata = self.draft_service.build_pile_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type="ai",
            original_text=instruction_value,
            ocr_text=recognized_text,
            input_text=recognized_text,
            last_source_filename=image_filename or "",
            pile_batches=[ai_batch],
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.piles)
        return self.get_draft_details(draft_id)

    def update_draft_pile_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "piles":
            raise ValueError("Черновик не является КП на сваи.")

        grade = (concrete_grade or "").strip()
        if not grade:
            raise ValueError("Укажите класс бетона.")

        order_data = list(payload.get("order_data") or [])
        if not order_data:
            raise ValueError("Список свай пустой.")

        lines: list[str] = []
        for item in order_data:
            mark = str(item.get("mark") or item.get("name") or "").strip()
            qty = int(item.get("qty") or 0)
            if mark and qty > 0:
                lines.append(f"{mark} {grade} {qty}")

        if not lines:
            raise ValueError("Список свай пустой.")

        next_text = "\n".join(lines)
        preview = self.pile_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_pile_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            pile_batches=list(metadata.get("pile_batches") or []),
            source_metadata={},
        )
        next_metadata["default_concrete_grade"] = grade
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.piles)
        return self.get_draft_details(draft_id)

    async def update_draft_marches(
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
        if str(metadata.get("product_type", "plates")).lower() != "marches":
            raise ValueError("Черновик не является КП на лестничные марши.")

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="marches",
        )
        current_text = str(metadata.get("input_text", "") or "")
        normalized_mode = mode.strip().lower()
        if normalized_mode not in {"append", "replace"}:
            raise ValueError("Некорректный режим обновления списка маршей.")
        if normalized_mode == "append" and current_text:
            next_text = self._merge_plate_texts(current_text, source_text["input_text"])
            batches = list(metadata.get("march_batches") or [])
            batches.append(source_text["batch"])
        else:
            next_text = source_text["input_text"]
            batches = [source_text["batch"]]

        preview = self.march_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_march_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=next_text,
            last_source_filename=source_text["filename"],
            march_batches=batches,
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.marches)
        return self.get_draft_details(draft_id)

    async def apply_ai_marches_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "marches":
            raise ValueError("ИИ-редактирование маршей доступно только для КП на лестничные марши.")

        instruction_value = (instruction or "").strip()
        if len(instruction_value) < 3:
            raise ValueError("Инструкция для ИИ должна содержать минимум 3 символа.")

        current_text = str(metadata.get("input_text", "") or "")
        tmp_path: Path | None = None
        try:
            if image_bytes:
                suffix = _safe_ocr_temp_suffix(image_filename)
                with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(image_bytes)
                    tmp_path = Path(tmp_file.name)
                result = await apply_marches_with_ai(
                    current_marches_text=current_text,
                    user_instruction=instruction_value,
                    image_path=str(tmp_path),
                )
            else:
                result = await apply_marches_with_ai(
                    current_marches_text=current_text,
                    user_instruction=instruction_value,
                )
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

        recognized_text = str((result or {}).get("text", "")).strip()
        if not recognized_text:
            raise ValueError("ИИ не смог обработать список маршей. Попробуйте уточнить инструкцию.")

        source_metadata = {
            "ai_applied": True,
            "last_ai_instruction": instruction_value,
            "ai_cost_usd": float((result or {}).get("cost_usd", 0.0) or 0.0),
            "ai_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ai_plates": list((result or {}).get("plates") or []),
            "ocr_plates": list((result or {}).get("plates") or []),
            "ocr_draft_plates": list((result or {}).get("draft_plates") or []),
            "ocr_corrections": list((result or {}).get("corrections") or []),
            "ocr_verify_applied": False,
            "ocr_verify_failed": False,
            "ocr_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ocr_row_count_on_image": (result or {}).get("row_count_on_image"),
        }
        ai_batch = {
            "source_type": "ai",
            "original_text": instruction_value,
            "normalized_text": recognized_text,
            "ocr_text": recognized_text,
            "filename": image_filename or "",
        }
        preview = self.march_service.generate_preview(
            recognized_text,
            db_path=str(DB_PATH),
        )
        next_metadata = self.draft_service.build_march_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type="ai",
            original_text=instruction_value,
            ocr_text=recognized_text,
            input_text=recognized_text,
            last_source_filename=image_filename or "",
            march_batches=[ai_batch],
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.marches)
        return self.get_draft_details(draft_id)

    def update_draft_march_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "marches":
            raise ValueError("Черновик не является КП на лестничные марши.")

        grade = (concrete_grade or "").strip()
        if not grade:
            raise ValueError("Укажите класс бетона.")

        order_data = list(payload.get("order_data") or [])
        if not order_data:
            raise ValueError("Список маршей пустой.")

        lines: list[str] = []
        for item in order_data:
            mark = str(item.get("mark") or item.get("name") or "").strip()
            qty = int(item.get("qty") or 0)
            if mark and qty > 0:
                lines.append(f"{mark} {grade} {qty}")

        if not lines:
            raise ValueError("Список маршей пустой.")

        next_text = "\n".join(lines)
        preview = self.march_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_march_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            march_batches=list(metadata.get("march_batches") or []),
            source_metadata={},
        )
        next_metadata["default_concrete_grade"] = grade
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.marches)
        return self.get_draft_details(draft_id)

    async def update_draft_bridge_piles(
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
        if str(metadata.get("product_type", "plates")).lower() != "bridge_piles":
            raise ValueError("Черновик не является КП на мостовые сваи.")

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="bridge_piles",
        )
        current_text = str(metadata.get("input_text", "") or "")
        normalized_mode = mode.strip().lower()
        if normalized_mode not in {"append", "replace"}:
            raise ValueError("Некорректный режим обновления списка мостовых свай.")
        if normalized_mode == "append" and current_text:
            next_text = self._merge_plate_texts(current_text, source_text["input_text"])
            batches = list(metadata.get("bridge_pile_batches") or [])
            batches.append(source_text["batch"])
        else:
            next_text = source_text["input_text"]
            batches = [source_text["batch"]]

        preview = self.bridge_pile_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_bridge_pile_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=next_text,
            last_source_filename=source_text["filename"],
            bridge_pile_batches=batches,
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.bridge_piles)
        return self.get_draft_details(draft_id)

    async def apply_ai_bridge_piles_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "bridge_piles":
            raise ValueError("ИИ-редактирование доступно только для КП на мостовые сваи.")

        instruction_value = (instruction or "").strip()
        if len(instruction_value) < 3:
            raise ValueError("Инструкция для ИИ должна содержать минимум 3 символа.")

        current_text = str(metadata.get("input_text", "") or "")
        tmp_path: Path | None = None
        try:
            if image_bytes:
                suffix = _safe_ocr_temp_suffix(image_filename)
                with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(image_bytes)
                    tmp_path = Path(tmp_file.name)
                result = await apply_bridge_piles_with_ai(
                    current_bridge_piles_text=current_text,
                    user_instruction=instruction_value,
                    image_path=str(tmp_path),
                )
            else:
                result = await apply_bridge_piles_with_ai(
                    current_bridge_piles_text=current_text,
                    user_instruction=instruction_value,
                )
        finally:
            if tmp_path is not None:
                try:
                    tmp_path.unlink(missing_ok=True)
                except OSError:
                    pass

        recognized_text = str(result.get("text") or "").strip()
        if not recognized_text:
            raise ValueError("ИИ не вернул распознанный список мостовых свай.")

        preview = self.bridge_pile_service.generate_preview(recognized_text, db_path=str(DB_PATH))
        ai_batch = {
            "source_type": "ai",
            "filename": image_filename or "",
            "text": recognized_text,
        }
        source_metadata = {
            "ai_applied": True,
            "ocr_warnings": list(result.get("warnings") or []),
        }
        next_metadata = self.draft_service.build_bridge_pile_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type="ai",
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=recognized_text,
            input_text=recognized_text,
            last_source_filename=image_filename or "",
            bridge_pile_batches=[ai_batch],
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.bridge_piles)
        return self.get_draft_details(draft_id)

    def update_draft_bridge_pile_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        """Apply grade where available; skip others + warning (decision A / Q12)."""
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "bridge_piles":
            raise ValueError("Черновик не является КП на мостовые сваи.")

        grade = (concrete_grade or "").strip()
        if not grade:
            raise ValueError("Укажите класс бетона.")

        order_data = list(payload.get("order_data") or [])
        if not order_data:
            raise ValueError("Список мостовых свай пустой.")

        lines: list[str] = []
        skipped: list[str] = []
        for item in order_data:
            mark = str(item.get("mark") or item.get("name") or "").strip()
            qty = int(item.get("qty") or 0)
            if not mark or qty <= 0:
                continue
            available = list(item.get("available_grades") or []) or list_available_grades(
                mark, db_path=str(DB_PATH)
            )
            if grade in available:
                lines.append(f"{mark} {grade} {qty}")
            else:
                current_grade = str(item.get("concrete_grade") or "B25").strip()
                lines.append(f"{mark} {current_grade} {qty}")
                skipped.append(mark)

        if not lines:
            raise ValueError("Список мостовых свай пустой.")

        next_text = "\n".join(lines)
        preview = self.bridge_pile_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_bridge_pile_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            bridge_pile_batches=list(metadata.get("bridge_pile_batches") or []),
            source_metadata={},
        )
        next_metadata["default_concrete_grade"] = grade
        if skipped:
            warning = (
                f"Класс {grade} не применён к маркам без цены в прайсе: "
                + ", ".join(skipped)
            )
            warnings = list(next_metadata.get("warnings") or [])
            if warning not in warnings:
                warnings.append(warning)
            next_metadata["warnings"] = warnings
            next_metadata["grade_bulk_skipped_marks"] = skipped
        else:
            next_metadata.pop("grade_bulk_skipped_marks", None)

        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.bridge_piles)
        return self.get_draft_details(draft_id)

    async def _create_fbs_draft(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
    ) -> dict[str, Any]:
        text_value = (text or "").strip()
        if not text_value and not image_bytes:
            metadata = {
                "product_type": "fbs",
                "owner_user_id": owner_user_id,
                "current_step": WizardStepId.fbs.value,
                "wide_plates_resolved": True,
                "default_concrete_grade": "B25",
            }
            draft_id = self.draft_store.save_preview(
                order=PlateOrder(),
                optimization_context=OptimizationContext(order=PlateOrder()),
                order_data=[],
                metadata=metadata,
            )
            return self.get_draft_details(draft_id)

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="fbs",
        )
        preview = self.fbs_service.generate_preview(
            source_text["input_text"],
            db_path=str(DB_PATH),
        )
        batches = [source_text["batch"]]
        metadata = self.draft_service.build_fbs_preview_metadata(
            preview=preview,
            base_metadata={"product_type": "fbs"},
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=source_text["input_text"],
            last_source_filename=source_text["filename"],
            fbs_batches=batches,
            source_metadata=source_metadata,
            owner_user_id=owner_user_id,
        )
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.fbs)
        return self.get_draft_details(draft_id)

    async def update_draft_fbs(
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
        if str(metadata.get("product_type", "plates")).lower() != "fbs":
            raise ValueError("Черновик не является КП на ФБС.")

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="fbs",
        )
        current_text = str(metadata.get("input_text", "") or "")
        normalized_mode = mode.strip().lower()
        if normalized_mode not in {"append", "replace"}:
            raise ValueError("Некорректный режим обновления списка ФБС.")
        if normalized_mode == "append" and current_text:
            next_text = self._merge_plate_texts(current_text, source_text["input_text"])
            batches = list(metadata.get("fbs_batches") or [])
            batches.append(source_text["batch"])
        else:
            next_text = source_text["input_text"]
            batches = [source_text["batch"]]

        preview = self.fbs_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_fbs_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=next_text,
            last_source_filename=source_text["filename"],
            fbs_batches=batches,
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.fbs)
        return self.get_draft_details(draft_id)

    async def apply_ai_fbs_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "fbs":
            raise ValueError("ИИ-редактирование доступно только для КП на ФБС.")

        instruction_value = (instruction or "").strip()
        if len(instruction_value) < 3:
            raise ValueError("Инструкция для ИИ должна содержать минимум 3 символа.")

        current_text = str(metadata.get("input_text", "") or "")
        tmp_path: Path | None = None
        try:
            if image_bytes:
                suffix = _safe_ocr_temp_suffix(image_filename)
                with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(image_bytes)
                    tmp_path = Path(tmp_file.name)
                result = await apply_fbs_with_ai(
                    current_fbs_text=current_text,
                    user_instruction=instruction_value,
                    image_path=str(tmp_path),
                )
            else:
                result = await apply_fbs_with_ai(
                    current_fbs_text=current_text,
                    user_instruction=instruction_value,
                )
        finally:
            if tmp_path is not None:
                try:
                    tmp_path.unlink(missing_ok=True)
                except OSError:
                    pass

        recognized_text = str(result.get("text") or "").strip()
        if not recognized_text:
            raise ValueError("ИИ не вернул распознанный список ФБС.")

        preview = self.fbs_service.generate_preview(recognized_text, db_path=str(DB_PATH))
        ai_batch = {
            "source_type": "ai",
            "filename": image_filename or "",
            "text": recognized_text,
        }
        source_metadata = {
            "ai_applied": True,
            "ocr_warnings": list(result.get("warnings") or []),
        }
        next_metadata = self.draft_service.build_fbs_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type="ai",
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=recognized_text,
            input_text=recognized_text,
            last_source_filename=image_filename or "",
            fbs_batches=[ai_batch],
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.fbs)
        return self.get_draft_details(draft_id)

    def update_draft_fbs_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        """Apply grade where available; skip others + warning (decision A / Q12)."""
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "fbs":
            raise ValueError("Черновик не является КП на ФБС.")

        grade = (concrete_grade or "").strip()
        if not grade:
            raise ValueError("Укажите класс бетона.")

        order_data = list(payload.get("order_data") or [])
        if not order_data:
            raise ValueError("Список ФБС пустой.")

        lines: list[str] = []
        skipped: list[str] = []
        for item in order_data:
            mark = str(item.get("mark") or item.get("name") or "").strip()
            qty = int(item.get("qty") or 0)
            if not mark or qty <= 0:
                continue
            available = list(item.get("available_grades") or []) or list_fbs_available_grades(
                mark, db_path=str(DB_PATH)
            )
            if grade in available:
                lines.append(f"{mark} {grade} {qty}")
            else:
                current_grade = str(item.get("concrete_grade") or "B25").strip()
                lines.append(f"{mark} {current_grade} {qty}")
                skipped.append(mark)

        if not lines:
            raise ValueError("Список ФБС пустой.")

        next_text = "\n".join(lines)
        preview = self.fbs_service.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_fbs_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            fbs_batches=list(metadata.get("fbs_batches") or []),
            source_metadata={},
        )
        next_metadata["default_concrete_grade"] = grade
        if skipped:
            warning = (
                f"Класс {grade} не применён к маркам без цены в прайсе: "
                + ", ".join(skipped)
            )
            warnings = list(next_metadata.get("warnings") or [])
            if warning not in warnings:
                warnings.append(warning)
            next_metadata["warnings"] = warnings
            next_metadata["grade_bulk_skipped_marks"] = skipped
        else:
            next_metadata.pop("grade_bulk_skipped_marks", None)

        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.fbs)
        return self.get_draft_details(draft_id)

    async def update_draft_steps(
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
        if str(metadata.get("product_type", "plates")).lower() != "steps":
            raise ValueError("Черновик не является КП на ступени.")

        source_text, source_metadata = await self.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type="steps",
        )
        current_text = str(metadata.get("input_text", "") or "")
        normalized_mode = mode.strip().lower()
        if normalized_mode not in {"append", "replace"}:
            raise ValueError("Некорректный режим обновления списка ступеней.")
        if normalized_mode == "append" and current_text:
            next_text = self._merge_plate_texts(current_text, source_text["input_text"])
            batches = list(metadata.get("step_batches") or [])
            batches.append(source_text["batch"])
        else:
            next_text = source_text["input_text"]
            batches = [source_text["batch"]]

        preview = self.step_service_product.generate_preview(next_text, db_path=str(DB_PATH))
        next_metadata = self.draft_service.build_step_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=next_text,
            last_source_filename=source_text["filename"],
            step_batches=batches,
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.steps)
        return self.get_draft_details(draft_id)

    async def apply_ai_steps_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != "steps":
            raise ValueError("ИИ-редактирование ступеней доступно только для КП на ступени.")

        instruction_value = (instruction or "").strip()
        if len(instruction_value) < 3:
            raise ValueError("Инструкция для ИИ должна содержать минимум 3 символа.")

        current_text = str(metadata.get("input_text", "") or "")
        tmp_path: Path | None = None
        try:
            if image_bytes:
                suffix = _safe_ocr_temp_suffix(image_filename)
                with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(image_bytes)
                    tmp_path = Path(tmp_file.name)
                result = await apply_steps_with_ai(
                    current_steps_text=current_text,
                    user_instruction=instruction_value,
                    image_path=str(tmp_path),
                )
            else:
                result = await apply_steps_with_ai(
                    current_steps_text=current_text,
                    user_instruction=instruction_value,
                )
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

        recognized_text = str((result or {}).get("text", "")).strip()
        if not recognized_text:
            raise ValueError("ИИ не смог обработать список ступеней. Попробуйте уточнить инструкцию.")

        source_metadata = {
            "ai_applied": True,
            "last_ai_instruction": instruction_value,
            "ai_cost_usd": float((result or {}).get("cost_usd", 0.0) or 0.0),
            "ai_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ai_plates": list((result or {}).get("plates") or []),
            "ocr_plates": list((result or {}).get("plates") or []),
            "ocr_draft_plates": list((result or {}).get("draft_plates") or []),
            "ocr_corrections": list((result or {}).get("corrections") or []),
            "ocr_verify_applied": False,
            "ocr_verify_failed": False,
            "ocr_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ocr_row_count_on_image": (result or {}).get("row_count_on_image"),
        }
        ai_batch = {
            "source_type": "ai",
            "original_text": instruction_value,
            "normalized_text": recognized_text,
            "ocr_text": recognized_text,
            "filename": image_filename or "",
        }
        preview = self.step_service_product.generate_preview(
            recognized_text,
            db_path=str(DB_PATH),
        )
        next_metadata = self.draft_service.build_step_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type="ai",
            original_text=instruction_value,
            ocr_text=recognized_text,
            input_text=recognized_text,
            last_source_filename=image_filename or "",
            step_batches=[ai_batch],
            source_metadata=source_metadata,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=preview.order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.steps)
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
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        draft = await self.create_draft(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            owner_user_id=owner_user_id,
            plate_order_ctx=plate_order_ctx,
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
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        product_type = str(metadata.get("product_type", "plates")).lower()
        if product_type == "piles":
            raise ValueError("Для КП на сваи используйте endpoint /piles.")
        if product_type == "marches":
            raise ValueError("Для КП на лестничные марши используйте endpoint /marches.")
        if product_type == "steps":
            raise ValueError("Для КП на ступени используйте endpoint /steps.")
        source_text, source_metadata = await self.draft_service.resolve_source_input(
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

        preview = self.commercial_service.generate_preview(
            text=next_text,
            plate_order_ctx=plate_order_ctx,
        )
        next_metadata = self.draft_service.build_preview_metadata(
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

    async def apply_ai_plates_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        instruction_value = (instruction or "").strip()
        if len(instruction_value) < 3:
            raise ValueError("Инструкция для ИИ должна содержать минимум 3 символа.")

        current_text = str(metadata.get("input_text", "") or "")
        tmp_path: Path | None = None
        try:
            if image_bytes:
                suffix = _safe_ocr_temp_suffix(image_filename)
                with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(image_bytes)
                    tmp_path = Path(tmp_file.name)
                result = await apply_plates_with_ai(
                    current_plates_text=current_text,
                    user_instruction=instruction_value,
                    image_path=str(tmp_path),
                )
            else:
                result = await apply_plates_with_ai(
                    current_plates_text=current_text,
                    user_instruction=instruction_value,
                )
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)

        recognized_text = str((result or {}).get("text", "")).strip()
        if not recognized_text:
            raise ValueError("ИИ не смог обработать список плит. Попробуйте уточнить инструкцию.")

        source_metadata = {
            "ai_applied": True,
            "last_ai_instruction": instruction_value,
            "ai_cost_usd": float((result or {}).get("cost_usd", 0.0) or 0.0),
            "ai_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ai_plates": list((result or {}).get("plates") or []),
            "ocr_plates": list((result or {}).get("plates") or []),
            "ocr_draft_plates": list((result or {}).get("draft_plates") or []),
            "ocr_corrections": list((result or {}).get("corrections") or []),
            "ocr_verify_applied": False,
            "ocr_verify_failed": False,
            "ocr_method": str((result or {}).get("method") or "GPT-4o+ai"),
            "ocr_row_count_on_image": (result or {}).get("row_count_on_image"),
        }
        ai_batch = {
            "source_type": "ai",
            "original_text": instruction_value,
            "normalized_text": recognized_text,
            "ocr_text": recognized_text,
            "filename": image_filename or "",
        }
        preview = self.commercial_service.generate_preview(
            text=recognized_text,
            plate_order_ctx=plate_order_ctx,
        )
        next_metadata = self.draft_service.build_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type="ai",
            original_text=instruction_value,
            ocr_text=recognized_text,
            input_text=recognized_text,
            last_source_filename=image_filename or "",
            plate_batches=[ai_batch],
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

    def resolve_wide_plates(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
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
                replacement_lines = self._normalize_replacement_lines(
                    replacement_text,
                    plate_order_ctx=plate_order_ctx,
                )
                if not replacement_lines:
                    raise ValueError("Для замены широкой плиты нужно указать корректный список замен.")
                merged_lines.extend(replacement_lines)
                continue
            raise ValueError("Некорректное действие для обработки широкой плиты.")

        if not merged_lines:
            raise ValueError("После обработки широких плит список стал пустым.")

        next_text = "\n".join(merged_lines)
        plate_batches = list(metadata.get("plate_batches") or [])
        updated_batches = self._apply_wide_plate_decisions_to_batches(
            plate_batches,
            wide_set,
            resolved_by_line,
            plate_order_ctx=plate_order_ctx,
        )
        preview = self.commercial_service.generate_preview(
            text=next_text,
            plate_order_ctx=plate_order_ctx,
        )
        next_metadata = self.draft_service.build_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            plate_batches=updated_batches,
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
        self._persist_wizard_step(draft_id, WizardStepId.plates)
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
            else:
                self._persist_wizard_step(draft_id, WizardStepId.client)

        return self.get_draft_details(draft_id)

    def calculate_draft(self, draft_id: str) -> dict[str, Any]:
        details = self.get_draft_details(draft_id)
        metadata = details["metadata"]
        self.calculation_service.enforce_calculate_prerequisites(
            order_data=list(details["order_data"]),
            metadata=dict(metadata),
        )
        self.draft_store.update_metadata(draft_id, current_step=WizardStepId.result.value)
        return self.get_draft_details(draft_id)

    def get_draft_breakdown(self, draft_id: str) -> dict[str, Any]:
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        raw_tables = metadata.get("breakdown_tables") or []
        items: list[dict[str, Any]] = []
        for table in raw_tables:
            if not isinstance(table, dict):
                continue
            name = str(table.get("name", "") or "").strip()
            rows_raw = table.get("rows") or []
            rows: list[list[str]] = []
            for row in rows_raw:
                if isinstance(row, (list, tuple)):
                    cells = [str(cell) for cell in row[:3]]
                    while len(cells) < 3:
                        cells.append("")
                    rows.append(cells)
            items.append({"name": name, "rows": rows})
        return {"draft_id": draft_id, "items": items}

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
            if key not in ("breakdown_tables", "owner_user_id", "schema_file")
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
            "files": self.export_service.collect_draft_files(metadata, draft_id),
            "saved_offer": self._normalize_saved_offer(metadata.get("saved_offer")),
            "totals": totals,
            "offer_identity": self.export_service.build_offer_identity_payload(draft_id),
        }

    def generate_files(
        self,
        draft_id: str,
        file_types: Iterable[str] | None = None,
        *,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> list[dict[str, str]]:
        payload = self._load_draft_or_raise(draft_id)
        return self.export_service.generate_files(
            draft_id,
            payload,
            file_types,
            plate_order_ctx=plate_order_ctx,
        )

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
            resolved = self.export_service.resolve_generated_file(xlsx_file["filename"])
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
            product_type=str(metadata.get("product_type", "plates") or "plates"),
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
        offer_identity = self.export_service.build_offer_identity_payload(draft_id)
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

    def _load_draft_or_raise(self, draft_id: str) -> dict[str, Any]:
        try:
            payload = self.draft_store.load_preview(draft_id)
        except UnsafeDraftIdError:
            raise FileNotFoundError(f"Draft '{draft_id}' not found.") from None
        if payload is None:
            raise FileNotFoundError(f"Draft '{draft_id}' not found.")
        return payload

    def _build_preview_metadata(self, **kwargs: Any) -> dict[str, Any]:
        return self.draft_service.build_preview_metadata(**kwargs)

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

    def _normalize_wide_plate_lines(self, items: Iterable[Any]) -> list[dict[str, Any]]:
        return self.draft_service.serialize_wide_plate_lines(items)

    def _merge_plate_texts(self, current_text: str, next_text: str) -> str:
        parts = [current_text.strip(), next_text.strip()]
        return "\n".join(part for part in parts if part)

    def _normalize_replacement_lines(
        self,
        replacement_text: str,
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> list[str]:
        if not replacement_text.strip():
            return []
        preview = self.commercial_service.generate_preview(
            text=replacement_text,
            plate_order_ctx=plate_order_ctx,
        )
        return list(preview.parse_result.normalized_lines)

    def _apply_wide_plate_decisions_to_batches(
        self,
        plate_batches: list[dict[str, Any]],
        wide_set: set[str],
        resolved_by_line: dict[str, dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> list[dict[str, Any]]:
        if not plate_batches:
            return plate_batches

        updated: list[dict[str, Any]] = []
        for batch in plate_batches:
            batch_text = str(batch.get("normalized_text", "") or "")
            batch_lines = [line.strip() for line in batch_text.split("\n") if line.strip()]
            next_lines: list[str] = []
            for line in batch_lines:
                if line not in wide_set:
                    next_lines.append(line)
                    continue
                decision = resolved_by_line.get(line)
                if decision is None:
                    next_lines.append(line)
                    continue
                action = str(decision.get("action", "")).strip().lower()
                if action == "confirm":
                    next_lines.append(line)
                elif action == "exclude":
                    continue
                elif action == "replace":
                    replacement_text = str(decision.get("replacement_text", "") or "").strip()
                    replacement_lines = self._normalize_replacement_lines(
                        replacement_text,
                        plate_order_ctx=plate_order_ctx,
                    )
                    next_lines.extend(replacement_lines)
                else:
                    next_lines.append(line)
            next_batch = dict(batch)
            next_batch["normalized_text"] = "\n".join(next_lines)
            updated.append(next_batch)
        return updated

    def get_or_generate_file(self, safe_filename: str) -> Path:
        """Path under configured ``outputs_dir`` for a generated file basename (no subpaths)."""
        return self.export_service.get_or_generate_file(safe_filename)

    def _resolve_generated_file(self, filename: str) -> Path:
        return self.export_service.resolve_generated_file(filename)
