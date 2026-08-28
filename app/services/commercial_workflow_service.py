from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable
from uuid import uuid4

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
from core.kp_order_data import order_data_from_kp_info
from core.ocr_gpt import (
    apply_bridge_piles_with_ai,
    apply_fbs_with_ai,
    apply_marches_with_ai,
    apply_piles_with_ai,
    apply_plates_with_ai,
    apply_steps_with_ai,
)
from core.plate_order_context import PlateOrderContext

_APPEND_PRODUCT_TYPES = frozenset(
    {"plates", "piles", "steps", "marches", "bridge_piles", "fbs"}
)
_PRODUCT_TYPE_TO_WIZARD_STEP = {
    "plates": WizardStepId.plates,
    "piles": WizardStepId.piles,
    "steps": WizardStepId.steps,
    "marches": WizardStepId.marches,
    "bridge_piles": WizardStepId.bridge_piles,
    "fbs": WizardStepId.fbs,
}


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

    def _stamp_order_data(
        self,
        order_data: list[Any] | None,
        *,
        product_type: str,
        previous_order_data: list[Any] | None = None,
    ) -> list[dict[str, Any]]:
        return self.draft_service.stamp_order_line_identity(
            list(order_data or []),
            product_type=product_type,
            previous_order_data=list(previous_order_data or []) if previous_order_data else None,
        )

    _PRODUCT_KIND_TO_TYPE: dict[str, str] = {
        "march": "marches",
        "step": "steps",
        "pile": "piles",
        "bridge_pile": "bridge_piles",
        "fbs": "fbs",
        "plate": "plates",
    }

    @staticmethod
    def _line_product_type(line: dict[str, Any] | None) -> str:
        if not isinstance(line, dict):
            return ""
        explicit = str(line.get("product_type") or "").strip().lower()
        if explicit:
            return explicit
        kind = str(line.get("product_kind") or "").strip().lower()
        return CommercialWorkflowService._PRODUCT_KIND_TO_TYPE.get(kind, "")

    @staticmethod
    def _line_is_sealed(line: dict[str, Any] | None) -> bool:
        """Sealed lines carry append_batch_id (assigned by ``_seal_unbatched_lines``)."""
        if not isinstance(line, dict):
            return False
        return bool(str(line.get("append_batch_id") or "").strip())

    def _partition_order_by_product_type(
        self,
        order_data: list[Any] | None,
        *,
        product_type: str,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        """Split order into (other types kept, same product_type lines).

        Untyped legacy mono lines (no product_type / resolvable product_kind) stay in
        ``same`` when the order has no conflicting typed product — so replace/bulk
        grade does not duplicate create-time rows that predate stamp_order_line_identity.
        """
        normalized = (product_type or "").strip().lower()
        rows = [dict(raw) for raw in list(order_data or []) if isinstance(raw, dict)]
        typed_conflict = any(
            (t := self._line_product_type(line)) and t != normalized for line in rows
        )
        others: list[dict[str, Any]] = []
        same: list[dict[str, Any]] = []
        for line in rows:
            line_type = self._line_product_type(line)
            if line_type == normalized or (not line_type and not typed_conflict):
                same.append(line)
            else:
                others.append(line)
        return others, same

    def _stamp_previous_for_product_update(
        self,
        same_previous: list[dict[str, Any]],
        *,
        mode: str,
        merged_cycle_text: bool,
    ) -> list[dict[str, Any]]:
        """Pick prior same-type lines whose line_ids may be reused when stamping.

        Append+merged cycle text reuses only unsealed same-type (current cycle).
        Replace reuses all same-type. Fresh append cycle reuses none.
        """
        normalized_mode = (mode or "").strip().lower()
        if merged_cycle_text:
            return [ln for ln in same_previous if not self._line_is_sealed(ln)]
        if normalized_mode == "replace":
            return list(same_previous)
        return []

    def _compose_order_data_for_product_update(
        self,
        *,
        previous_order_data: list[Any] | None,
        new_type_lines: list[dict[str, Any]],
        product_type: str,
        mode: str,
        merged_cycle_text: bool,
    ) -> list[dict[str, Any]]:
        """Keep other product types; compose same-type lines for append/replace.

        Append with cleared cycle input preserves chronological order (full previous
        list + new lines). Append with merged cycle text keeps sealed lines of any
        type chronologically and replaces only unsealed same-type lines.
        Replace keeps other types + new same-type.
        """
        normalized_mode = (mode or "").strip().lower()
        normalized_type = (product_type or "").strip().lower()
        if normalized_mode == "append" and not merged_cycle_text:
            return list(previous_order_data or []) + list(new_type_lines)
        if normalized_mode == "append" and merged_cycle_text:
            kept: list[dict[str, Any]] = []
            for raw in list(previous_order_data or []):
                if not isinstance(raw, dict):
                    continue
                line = dict(raw)
                unsealed_same_type = (
                    self._line_product_type(line) == normalized_type
                    and not self._line_is_sealed(line)
                )
                if not unsealed_same_type:
                    kept.append(line)
            return kept + list(new_type_lines)
        others, _same_previous = self._partition_order_by_product_type(
            previous_order_data,
            product_type=product_type,
        )
        return others + list(new_type_lines)

    def _current_cycle_lines(
        self,
        order_data: list[Any] | None,
        *,
        product_type: str,
    ) -> list[dict[str, Any]]:
        """Unsealed lines of product_type for the in-progress append/input cycle.

        Sealed lines (append_batch_id set) belong to prior batches and must not
        appear in input-step grade edits or cycle text rebuilds.
        """
        normalized = (product_type or "").strip().lower()
        out: list[dict[str, Any]] = []
        for raw in list(order_data or []):
            if not isinstance(raw, dict):
                continue
            if self._line_is_sealed(raw):
                continue
            line_type = self._line_product_type(raw)
            if not line_type or line_type == normalized:
                # Untyped legacy mono: include only when nothing in order is sealed
                # and no conflicting typed lines exist.
                if not line_type:
                    continue
                out.append(dict(raw))
        if out:
            return out
        unsealed = [
            dict(raw)
            for raw in list(order_data or [])
            if isinstance(raw, dict) and not self._line_is_sealed(raw)
        ]
        if unsealed and all(not self._line_product_type(line) for line in unsealed):
            return unsealed
        return []

    def _seal_unbatched_lines(
        self,
        order_data: list[Any] | None,
        metadata: dict[str, Any],
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        """Assign append_batch_id to unsealed lines; append metadata.append_batches entry."""
        lines = [dict(raw) for raw in list(order_data or []) if isinstance(raw, dict)]
        batches = [
            {
                "batch_id": str(batch.get("batch_id") or "").strip(),
                "product_type": str(batch.get("product_type") or "").strip().lower(),
                "line_ids": [str(lid) for lid in list(batch.get("line_ids") or []) if str(lid).strip()],
            }
            for batch in list(metadata.get("append_batches") or [])
            if isinstance(batch, dict) and str(batch.get("batch_id") or "").strip()
        ]

        unsealed_ids: list[str] = []
        for line in lines:
            if str(line.get("append_batch_id") or "").strip():
                continue
            line_id = str(line.get("line_id") or "").strip()
            if not line_id:
                continue
            unsealed_ids.append(line_id)

        if not unsealed_ids:
            return lines, batches

        product_type = str(metadata.get("product_type") or "plates").strip().lower() or "plates"
        if product_type not in _APPEND_PRODUCT_TYPES:
            product_type = "plates"
        batch_id = uuid4().hex
        unsealed_set = set(unsealed_ids)
        for line in lines:
            line_id = str(line.get("line_id") or "").strip()
            if line_id in unsealed_set:
                line["append_batch_id"] = batch_id
        batches.append(
            {
                "batch_id": batch_id,
                "product_type": product_type,
                "line_ids": unsealed_ids,
            }
        )
        return lines, batches

    def _persist_order_and_metadata(
        self,
        draft_id: str,
        *,
        payload: dict[str, Any],
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any],
    ) -> None:
        self.draft_store.replace_preview(
            draft_id,
            order=payload["order"],
            optimization_context=payload["optimization_context"],
            order_data=order_data,
            metadata=metadata,
        )

    def start_append_cycle(self, draft_id: str, *, product_type: str) -> dict[str, Any]:
        """Switch cycle product_type, clear cycle input, keep header + prior order_data."""
        normalized = (product_type or "").strip().lower()
        if normalized not in _APPEND_PRODUCT_TYPES:
            raise ValueError("Некорректный тип продукта.")

        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata") or {})
        order_data = list(payload.get("order_data") or [])

        sealed_order, sealed_batches = self._seal_unbatched_lines(order_data, metadata)
        metadata["append_batches"] = sealed_batches
        metadata["product_type"] = normalized
        metadata["input_text"] = ""
        metadata["original_text"] = ""
        metadata["ocr_text"] = ""
        metadata["normalized_text"] = ""
        metadata["normalized_lines"] = []
        metadata["accumulated_text"] = ""
        metadata["plate_batches"] = []
        metadata["pile_batches"] = []
        metadata["step_batches"] = []
        metadata["march_batches"] = []
        metadata["wide_plate_lines"] = []
        metadata["wide_plates_resolved"] = True
        metadata["unparsed_lines"] = []
        metadata["warnings"] = []
        metadata["last_source_filename"] = ""
        metadata["source_type"] = None

        wizard_step = _PRODUCT_TYPE_TO_WIZARD_STEP.get(normalized, WizardStepId.plates)
        metadata["current_step"] = wizard_step.value

        self._persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=sealed_order,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def undo_last_append_batch(self, draft_id: str) -> dict[str, Any]:
        """Remove the last append_batches entry and its lines from order_data."""
        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata") or {})
        batches = list(metadata.get("append_batches") or [])
        if not batches:
            raise ValueError("Нет циклов append для отмены.")

        last = batches[-1]
        if not isinstance(last, dict):
            raise ValueError("Нет циклов append для отмены.")
        remove_ids = {
            str(lid).strip()
            for lid in list(last.get("line_ids") or [])
            if str(lid).strip()
        }
        order_data = [
            dict(line)
            for line in list(payload.get("order_data") or [])
            if isinstance(line, dict)
            and str(line.get("line_id") or "").strip() not in remove_ids
        ]
        metadata["append_batches"] = batches[:-1]
        self._persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=order_data,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def delete_order_line(self, draft_id: str, line_id: str) -> dict[str, Any]:
        """Remove one order line and scrub its id from append_batches.line_ids."""
        target_id = (line_id or "").strip()
        if not target_id:
            raise ValueError("Не указан идентификатор строки.")

        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata") or {})
        order_data = [
            dict(line)
            for line in list(payload.get("order_data") or [])
            if isinstance(line, dict)
        ]
        if not any(str(line.get("line_id") or "").strip() == target_id for line in order_data):
            raise FileNotFoundError("Строка не найдена.")

        next_order = [
            line for line in order_data if str(line.get("line_id") or "").strip() != target_id
        ]
        next_batches: list[dict[str, Any]] = []
        for batch in list(metadata.get("append_batches") or []):
            if not isinstance(batch, dict):
                continue
            next_batch = dict(batch)
            next_batch["line_ids"] = [
                str(lid)
                for lid in list(batch.get("line_ids") or [])
                if str(lid).strip() and str(lid).strip() != target_id
            ]
            if next_batch["line_ids"]:
                next_batches.append(next_batch)
        metadata["append_batches"] = next_batches
        self._persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=next_order,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

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
        order_data = self._stamp_order_data(preview.order_data, product_type="plates")
        draft_id = self.draft_store.save_preview(
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="piles")
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="marches")
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="bridge_piles")
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="steps")
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        merged_cycle_text = bool(normalized_mode == "append" and current_text.strip())
        if merged_cycle_text:
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
        previous_order_data = list(payload.get("order_data") or [])
        _, same_previous = self._partition_order_by_product_type(
            previous_order_data,
            product_type="piles",
        )
        stamp_previous = self._stamp_previous_for_product_update(
            same_previous,
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        new_pile_lines = self._stamp_order_data(
            preview.order_data,
            product_type="piles",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_pile_lines,
            product_type="piles",
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="piles")
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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

        previous_order_data = list(payload.get("order_data") or [])
        cycle_items = self._current_cycle_lines(previous_order_data, product_type="piles")
        if not cycle_items:
            raise ValueError("Список свай пустой.")

        lines: list[str] = []
        for item in cycle_items:
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
        stamp_previous = self._stamp_previous_for_product_update(
            cycle_items,
            mode="append",
            merged_cycle_text=True,
        )
        new_pile_lines = self._stamp_order_data(
            preview.order_data,
            product_type="piles",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_pile_lines,
            product_type="piles",
            mode="append",
            merged_cycle_text=True,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        merged_cycle_text = bool(normalized_mode == "append" and current_text.strip())
        if merged_cycle_text:
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
        previous_order_data = list(payload.get("order_data") or [])
        _, same_previous = self._partition_order_by_product_type(
            previous_order_data,
            product_type="marches",
        )
        stamp_previous = self._stamp_previous_for_product_update(
            same_previous,
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        new_march_lines = self._stamp_order_data(
            preview.order_data,
            product_type="marches",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_march_lines,
            product_type="marches",
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="marches")
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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

        previous_order_data = list(payload.get("order_data") or [])
        cycle_items = self._current_cycle_lines(previous_order_data, product_type="marches")
        if not cycle_items:
            raise ValueError("Список маршей пустой.")

        lines: list[str] = []
        for item in cycle_items:
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
        stamp_previous = self._stamp_previous_for_product_update(
            cycle_items,
            mode="append",
            merged_cycle_text=True,
        )
        new_march_lines = self._stamp_order_data(
            preview.order_data,
            product_type="marches",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_march_lines,
            product_type="marches",
            mode="append",
            merged_cycle_text=True,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        merged_cycle_text = bool(normalized_mode == "append" and current_text.strip())
        if merged_cycle_text:
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
        previous_order_data = list(payload.get("order_data") or [])
        _, same_previous = self._partition_order_by_product_type(
            previous_order_data,
            product_type="bridge_piles",
        )
        stamp_previous = self._stamp_previous_for_product_update(
            same_previous,
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        new_bridge_pile_lines = self._stamp_order_data(
            preview.order_data,
            product_type="bridge_piles",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_bridge_pile_lines,
            product_type="bridge_piles",
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="bridge_piles")
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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

        previous_order_data = list(payload.get("order_data") or [])
        cycle_items = self._current_cycle_lines(previous_order_data, product_type="bridge_piles")
        if not cycle_items:
            raise ValueError("Список мостовых свай пустой.")

        lines: list[str] = []
        skipped: list[str] = []
        for item in cycle_items:
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

        stamp_previous = self._stamp_previous_for_product_update(
            cycle_items,
            mode="append",
            merged_cycle_text=True,
        )
        new_bridge_pile_lines = self._stamp_order_data(
            preview.order_data,
            product_type="bridge_piles",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_bridge_pile_lines,
            product_type="bridge_piles",
            mode="append",
            merged_cycle_text=True,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="fbs")
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        merged_cycle_text = bool(normalized_mode == "append" and current_text.strip())
        if merged_cycle_text:
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
        previous_order_data = list(payload.get("order_data") or [])
        _, same_previous = self._partition_order_by_product_type(
            previous_order_data,
            product_type="fbs",
        )
        stamp_previous = self._stamp_previous_for_product_update(
            same_previous,
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        new_fbs_lines = self._stamp_order_data(
            preview.order_data,
            product_type="fbs",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_fbs_lines,
            product_type="fbs",
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="fbs")
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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

        previous_order_data = list(payload.get("order_data") or [])
        cycle_items = self._current_cycle_lines(previous_order_data, product_type="fbs")
        if not cycle_items:
            raise ValueError("Список ФБС пустой.")

        lines: list[str] = []
        skipped: list[str] = []
        for item in cycle_items:
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

        stamp_previous = self._stamp_previous_for_product_update(
            cycle_items,
            mode="append",
            merged_cycle_text=True,
        )
        new_fbs_lines = self._stamp_order_data(
            preview.order_data,
            product_type="fbs",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_fbs_lines,
            product_type="fbs",
            mode="append",
            merged_cycle_text=True,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        merged_cycle_text = bool(normalized_mode == "append" and current_text.strip())
        if merged_cycle_text:
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
        previous_order_data = list(payload.get("order_data") or [])
        _, same_previous = self._partition_order_by_product_type(
            previous_order_data,
            product_type="steps",
        )
        stamp_previous = self._stamp_previous_for_product_update(
            same_previous,
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        new_step_lines = self._stamp_order_data(
            preview.order_data,
            product_type="steps",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_step_lines,
            product_type="steps",
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="steps")
        self.draft_store.replace_preview(
            draft_id,
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
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
        merged_cycle_text = bool(normalized_mode == "append" and current_text.strip())
        if merged_cycle_text:
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
        previous_order_data = list(payload.get("order_data") or [])
        _, same_previous = self._partition_order_by_product_type(
            previous_order_data,
            product_type="plates",
        )
        stamp_previous = self._stamp_previous_for_product_update(
            same_previous,
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        new_plate_lines = self._stamp_order_data(
            preview.order_data,
            product_type="plates",
            previous_order_data=stamp_previous,
        )
        order_data = self._compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_plate_lines,
            product_type="plates",
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        self.draft_store.replace_preview(
            draft_id,
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="plates")
        self.draft_store.replace_preview(
            draft_id,
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=order_data,
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
        order_data = self._stamp_order_data(preview.order_data, product_type="plates")
        self.draft_store.replace_preview(
            draft_id,
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=order_data,
            metadata=next_metadata,
        )
        self._persist_wizard_step(draft_id, WizardStepId.plates)
        return self.get_draft_details(draft_id)

    def resolve_unpriced_plates(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        from core.unpriced_plate_replacements import rewrite_plate_line_load

        payload = self._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        current_text = str(metadata.get("input_text", "") or "")
        if not current_text:
            raise ValueError("Список плит отсутствует.")

        unpriced_items = self.draft_service.serialize_unpriced_plate_lines(
            metadata.get("unpriced_plate_lines", [])
        )
        if not unpriced_items:
            return self.get_draft_details(draft_id)

        decisions_by_id: dict[str, dict[str, Any]] = {}
        decisions_by_line: dict[str, dict[str, Any]] = {}
        for item in decisions:
            line_id = str(item.get("line_id", "") or "").strip()
            source_line = str(item.get("source_line", "") or "").strip()
            if line_id:
                decisions_by_id[line_id] = item
            if source_line:
                decisions_by_line[source_line] = item

        resolved_by_line: dict[str, dict[str, Any]] = {}
        resolved_decisions: dict[str, dict[str, Any]] = {}
        unresolved: list[str] = []
        for unpriced_item in unpriced_items:
            line_id = str(unpriced_item.get("id", "")).strip()
            line = str(unpriced_item.get("line", "")).strip()
            decision = decisions_by_id.get(line_id) if line_id else None
            if decision is None:
                decision = decisions_by_line.get(line)
            if decision is None:
                unresolved.append(line or line_id)
                continue

            action = str(decision.get("action", "")).strip().lower()
            allowed_load_codes = {
                int(repl["load_code"])
                for repl in (unpriced_item.get("replacements") or [])
                if isinstance(repl, dict) and repl.get("load_code") is not None
            }
            if action == "replace_load":
                if not allowed_load_codes:
                    raise ValueError(
                        "Для позиции без производимых замен доступно только исключение."
                    )
                raw_load = decision.get("load_code")
                if raw_load is None:
                    raise ValueError("Для замены нагрузки нужно указать load_code.")
                try:
                    chosen_load = int(raw_load)
                except (TypeError, ValueError) as exc:
                    raise ValueError("Некорректный load_code для замены нагрузки.") from exc
                if chosen_load not in allowed_load_codes:
                    raise ValueError(
                        f"load_code={chosen_load} не входит в предложенные замены."
                    )
            elif action == "exclude":
                pass
            else:
                raise ValueError("Некорректное действие для позиции без цены.")

            resolved_decisions[line_id or line] = {
                "line_id": line_id or None,
                "source_line": line or None,
                "action": action,
                "load_code": decision.get("load_code"),
            }
            if line:
                resolved_by_line[line] = resolved_decisions[line_id or line]

        if unresolved:
            raise ValueError("Нужно выбрать действие для всех позиций без цены.")

        original_lines = [
            line.strip() for line in list(metadata.get("normalized_lines") or []) if line.strip()
        ]
        if not original_lines:
            original_lines = [
                line.strip() for line in re.split(r"[\n;]+", current_text) if line.strip()
            ]

        unpriced_set = {
            str(item.get("line", "")).strip()
            for item in unpriced_items
            if str(item.get("line", "")).strip()
        }

        merged_lines: list[str] = []
        for line in original_lines:
            matched_item = next(
                (
                    item
                    for item in unpriced_items
                    if str(item.get("line", "")).strip() == line
                    or (
                        str(item.get("name", "")).strip()
                        and str(item.get("name", "")).strip() in line
                    )
                ),
                None,
            )
            if matched_item is None:
                merged_lines.append(line)
                continue

            item_line = str(matched_item.get("line", "")).strip()
            item_id = str(matched_item.get("id", "")).strip()
            decision = (
                resolved_by_line.get(item_line)
                or resolved_decisions.get(item_id)
                or resolved_decisions.get(item_line)
            )
            if decision is None:
                merged_lines.append(line)
                continue

            action = str(decision.get("action", "")).strip().lower()
            if action == "exclude":
                continue
            if action == "replace_load":
                new_load = int(decision["load_code"])
                try:
                    merged_lines.append(rewrite_plate_line_load(line, new_load))
                except ValueError:
                    fallback = str(matched_item.get("name") or line)
                    qty_match = re.search(r"(\d+)\s*$", line.strip())
                    rewritten = rewrite_plate_line_load(fallback, new_load)
                    if qty_match and not re.search(r"\d+\s*$", rewritten.strip()):
                        rewritten = f"{rewritten} {qty_match.group(1)}"
                    merged_lines.append(rewritten)
                continue
            raise ValueError("Некорректное действие для позиции без цены.")

        if not merged_lines:
            raise ValueError("После обработки позиций без цены список стал пустым.")

        next_text = "\n".join(merged_lines)
        plate_batches = list(metadata.get("plate_batches") or [])
        updated_batches = self._apply_unpriced_plate_decisions_to_batches(
            plate_batches,
            unpriced_items,
            resolved_by_line,
            resolved_decisions,
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
            wide_plates_resolved=bool(metadata.get("wide_plates_resolved", True)),
            source_metadata={},
        )
        # Force resolved after explicit user action (mirror wide-plates).
        next_metadata["unpriced_plates_resolved"] = True
        next_metadata["unpriced_plate_decisions"] = list(resolved_decisions.values())
        order_data = self._stamp_order_data(preview.order_data, product_type="plates")
        self.draft_store.replace_preview(
            draft_id,
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=order_data,
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
        product_type = str(metadata.get("product_type", "plates") or "plates").lower()
        order_data = list(details["order_data"] or [])
        # Fill any missing identity (legacy drafts) without reminting existing ids.
        if any(
            not str((line or {}).get("line_id") or "").strip()
            or not str((line or {}).get("product_type") or "").strip()
            for line in order_data
            if isinstance(line, dict)
        ):
            stamped = self._stamp_order_data(
                order_data,
                product_type=product_type,
                previous_order_data=order_data,
            )
            payload = self._load_draft_or_raise(draft_id)
            self.draft_store.replace_preview(
                draft_id,
                order=payload["order"],
                optimization_context=payload["optimization_context"],
                order_data=stamped,
                metadata=dict(payload.get("metadata") or {}),
            )
            order_data = stamped
        self.calculation_service.enforce_calculate_prerequisites(
            order_data=order_data,
            metadata=dict(metadata),
        )
        payload = self._load_draft_or_raise(draft_id)
        meta = dict(payload.get("metadata") or {})
        sealed_order, sealed_batches = self._seal_unbatched_lines(
            list(payload.get("order_data") or order_data),
            meta,
        )
        meta["append_batches"] = sealed_batches
        meta["current_step"] = WizardStepId.result.value
        self._persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=sealed_order,
            metadata=meta,
        )
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

    def hydrate_draft_from_saved_kp(
        self,
        kp_id: int,
        *,
        owner_user_id: int,
    ) -> dict[str, Any]:
        """Create a draft bound to an existing KP for append (status «в работе» only)."""
        kp_raw = self.kp_repository.get_offer(kp_id)
        if not kp_raw:
            raise ValueError(f"КП №{kp_id} не найдено")

        status = str(kp_raw.get("status") or "").strip()
        if status != "в работе":
            raise ValueError("Дополнить КП можно только в статусе «в работе».")

        order_data = [
            dict(line)
            for line in order_data_from_kp_info(kp_raw)
            if isinstance(line, dict)
        ]
        cycle_type = "plates"
        for line in reversed(order_data):
            pt = str(line.get("product_type") or "").strip().lower()
            if pt in _APPEND_PRODUCT_TYPES:
                cycle_type = pt
                break

        manager_name = str(kp_raw.get("manager_name") or "").strip()
        manager_id, manager_phone, manager_email = self._resolve_manager_for_hydrate(
            manager_name
        )
        delivery = str(kp_raw.get("delivery_conditions") or "").strip()
        payment = str(kp_raw.get("payment_conditions") or "").strip()
        conditions_mode = "custom" if (delivery or payment) else "standard"
        execution_terms = str(kp_raw.get("execution_terms") or "")
        saved_offer = {
            "kp_id": int(kp_id),
            "status": status,
            "mode": "database",
            "execution_terms": execution_terms,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }
        metadata: dict[str, Any] = {
            "product_type": cycle_type,
            "owner_user_id": owner_user_id,
            "client_name": str(kp_raw.get("customer_name") or "").strip(),
            "manager_name": manager_name,
            "manager_id": manager_id,
            "manager_phone": manager_phone,
            "manager_email": manager_email,
            "discount_percent": float(kp_raw.get("discount_percent") or 0.0),
            "logistics_cost": float(kp_raw.get("logistics_cost") or 0.0),
            "delivery_conditions": delivery,
            "payment_conditions": payment,
            "conditions_mode": conditions_mode,
            "execution_terms": execution_terms,
            "wide_plates_resolved": True,
            "wide_plate_lines": [],
            "append_batches": [],
            "resume_kp_id": int(kp_id),
            "current_step": WizardStepId.result.value,
            "saved_offer": saved_offer,
        }
        draft_id = self.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def _resolve_manager_for_hydrate(
        self,
        manager_name: str,
    ) -> tuple[int | None, str, str]:
        """Match managers.fio; keep a truthy manager_id when only the name is known."""
        name = (manager_name or "").strip()
        if not name:
            return None, "", ""
        needle = name.casefold()
        for manager in self.manager_repository.list_managers():
            fio = str(manager.get("fio") or "").strip()
            if fio.casefold() != needle:
                continue
            return (
                int(manager["id"]),
                str(manager.get("contact_number") or ""),
                str(manager.get("email") or ""),
            )
        # Name known from KP but not in managers table — still sticky for result step.
        return 1, "", ""

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

        raw_owner = metadata.get("owner_user_id")
        owner_user_id = int(raw_owner) if raw_owner is not None else None
        customer_name = str(metadata.get("client_name", "") or "Клиент")
        manager_name = str(metadata.get("manager_name", "") or "")
        discount_percent = float(metadata.get("discount_percent", 0.0) or 0.0)
        logistics_cost = float(metadata.get("logistics_cost", 0.0) or 0.0)
        delivery_conditions = str(metadata.get("delivery_conditions", "") or "")
        payment_conditions = str(metadata.get("payment_conditions", "") or "")
        product_type = str(metadata.get("product_type", "plates") or "plates")
        order_data = payload["order_data"]

        # MNA-304 / Q1=C: resume append updates the same kp_id when draft is bound.
        existing_saved = payload.get("saved_offer") or metadata.get("saved_offer") or {}
        existing_kp_id = existing_saved.get("kp_id")
        if existing_kp_id is None:
            resume_kp_id = metadata.get("resume_kp_id")
            if resume_kp_id is not None:
                existing_kp_id = resume_kp_id
                existing_saved = {
                    **existing_saved,
                    "kp_id": resume_kp_id,
                    "status": existing_saved.get("status") or "в работе",
                }
        if existing_kp_id is not None:
            existing_status = str(existing_saved.get("status", "") or "").strip()
            if existing_status != "в работе":
                raise ValueError(
                    "Дополнить КП можно только в статусе «в работе»."
                )
            kp_id = self.kp_repository.update_offer_from_order_data(
                int(existing_kp_id),
                order_data=order_data,
                customer_name=customer_name,
                manager_name=manager_name,
                discount_percent=discount_percent,
                logistics_cost=logistics_cost,
                delivery_conditions=delivery_conditions,
                payment_conditions=payment_conditions,
                execution_terms=execution_terms,
                xlsx_path=xlsx_path,
                product_type=product_type,
            )
        else:
            kp_id = self.kp_repository.save_offer(
                creation_date=datetime.now().strftime("%d.%m.%Y"),
                customer_name=customer_name,
                manager_name=manager_name,
                discount_percent=discount_percent,
                logistics_cost=logistics_cost,
                delivery_conditions=delivery_conditions,
                payment_conditions=payment_conditions,
                execution_terms=execution_terms,
                status=status,
                order_data=order_data,
                xlsx_path=xlsx_path,
                owner_user_id=owner_user_id,
                product_type=product_type,
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

    def _apply_unpriced_plate_decisions_to_batches(
        self,
        plate_batches: list[dict[str, Any]],
        unpriced_items: list[dict[str, Any]],
        resolved_by_line: dict[str, dict[str, Any]],
        resolved_decisions: dict[str, dict[str, Any]],
    ) -> list[dict[str, Any]]:
        from core.unpriced_plate_replacements import rewrite_plate_line_load

        if not plate_batches:
            return plate_batches

        updated: list[dict[str, Any]] = []
        for batch in plate_batches:
            batch_text = str(batch.get("normalized_text", "") or "")
            batch_lines = [line.strip() for line in batch_text.split("\n") if line.strip()]
            next_lines: list[str] = []
            for line in batch_lines:
                matched_item = next(
                    (
                        item
                        for item in unpriced_items
                        if str(item.get("line", "")).strip() == line
                        or (
                            str(item.get("name", "")).strip()
                            and str(item.get("name", "")).strip() in line
                        )
                    ),
                    None,
                )
                if matched_item is None:
                    next_lines.append(line)
                    continue
                item_line = str(matched_item.get("line", "")).strip()
                item_id = str(matched_item.get("id", "")).strip()
                decision = (
                    resolved_by_line.get(item_line)
                    or resolved_decisions.get(item_id)
                    or resolved_decisions.get(item_line)
                )
                if decision is None:
                    next_lines.append(line)
                    continue
                action = str(decision.get("action", "")).strip().lower()
                if action == "exclude":
                    continue
                if action == "replace_load":
                    new_load = int(decision["load_code"])
                    try:
                        next_lines.append(rewrite_plate_line_load(line, new_load))
                    except ValueError:
                        fallback = str(matched_item.get("name") or line)
                        next_lines.append(rewrite_plate_line_load(fallback, new_load))
                    continue
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
