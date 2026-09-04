from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

from app.core.settings import get_settings
from app.schemas.commercial import WizardStepId
from app.repositories.kp_repository import KpRepository
from app.repositories.manager_repository import ManagerRepository
from app.services.commercial_calculation_service import CommercialCalculationService
from app.services.commercial_draft_service import CommercialDraftService
from app.services.commercial_draft_service import _safe_ocr_temp_suffix  # noqa: F401 — tests import from this module
from app.services.commercial_export_service import CommercialExportService
from app.services.commercial_bridge_pile_service import CommercialBridgePileService
from app.services.commercial_fbs_service import CommercialFbsService
from app.services.commercial_march_service import CommercialMarchService
from app.services.commercial_pile_service import CommercialPileService
from app.services.commercial_step_service import CommercialStepService
from app.services.commercial_service import CommercialService
from app.services.commercial_wizard_step_service import CommercialWizardStepService
from app.services.commercial_order_identity import (
    APPEND_PRODUCT_TYPES,
    CommercialOrderIdentity,
)
from app.services.commercial_draft_lifecycle import CommercialDraftLifecycle
from app.services.commercial_plate_resolve import CommercialPlateResolve
from app.services.product_draft_handler import ProductDraftHandler
from app.services.draft_store import DraftStore, UnsafeDraftIdError
from app.services.execution_terms_service import ExecutionTermsService
from core.commercial_offer_xlsx import DB_PATH  # noqa: F401 — tests monkeypatch this module
from core.commercial_pricing import ensure_order_priced  # noqa: F401 — tests monkeypatch this module
from core.ocr_gpt import apply_plates_with_ai  # re-export for product_draft_config / tests
from core.plate_order_context import PlateOrderContext

_APPEND_PRODUCT_TYPES = APPEND_PRODUCT_TYPES


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
        self.order_identity = CommercialOrderIdentity(self.draft_service)
        self.draft_lifecycle = CommercialDraftLifecycle(self)
        self.plate_resolve = CommercialPlateResolve(self)
        self.product_draft_handler = ProductDraftHandler(self)

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
        return self.order_identity.stamp_order_data(
            order_data,
            product_type=product_type,
            previous_order_data=previous_order_data,
        )

    def _line_product_type(self, line: dict[str, Any] | None) -> str:
        return self.order_identity.line_product_type(line)

    def _line_is_sealed(self, line: dict[str, Any] | None) -> bool:
        return self.order_identity.line_is_sealed(line)

    def _partition_order_by_product_type(
        self,
        order_data: list[Any] | None,
        *,
        product_type: str,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        return self.order_identity.partition_order_by_product_type(
            order_data,
            product_type=product_type,
        )

    def _stamp_previous_for_product_update(
        self,
        same_previous: list[dict[str, Any]],
        *,
        mode: str,
        merged_cycle_text: bool,
    ) -> list[dict[str, Any]]:
        return self.order_identity.stamp_previous_for_product_update(
            same_previous,
            mode=mode,
            merged_cycle_text=merged_cycle_text,
        )

    def _compose_order_data_for_product_update(
        self,
        *,
        previous_order_data: list[Any] | None,
        new_type_lines: list[dict[str, Any]],
        product_type: str,
        mode: str,
        merged_cycle_text: bool,
    ) -> list[dict[str, Any]]:
        return self.order_identity.compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_type_lines,
            product_type=product_type,
            mode=mode,
            merged_cycle_text=merged_cycle_text,
        )

    def _current_cycle_lines(
        self,
        order_data: list[Any] | None,
        *,
        product_type: str,
    ) -> list[dict[str, Any]]:
        return self.order_identity.current_cycle_lines(
            order_data,
            product_type=product_type,
        )

    def _seal_unbatched_lines(
        self,
        order_data: list[Any] | None,
        metadata: dict[str, Any],
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        return self.order_identity.seal_unbatched_lines(order_data, metadata)

    def _persist_order_and_metadata(
        self,
        draft_id: str,
        *,
        payload: dict[str, Any],
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any],
    ) -> None:
        self.draft_lifecycle.persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=order_data,
            metadata=metadata,
        )

    def start_append_cycle(self, draft_id: str, *, product_type: str) -> dict[str, Any]:
        """Switch cycle product_type, clear cycle input, keep header + prior order_data."""
        return self.draft_lifecycle.start_append_cycle(draft_id, product_type=product_type)

    def undo_last_append_batch(self, draft_id: str) -> dict[str, Any]:
        """Remove the last append_batches entry and its lines from order_data."""
        return self.draft_lifecycle.undo_last_append_batch(draft_id)

    def delete_order_line(self, draft_id: str, line_id: str) -> dict[str, Any]:
        """Remove one order line and scrub its id from append_batches.line_ids."""
        return self.draft_lifecycle.delete_order_line(draft_id, line_id)

    def patch_order_line(
        self,
        draft_id: str,
        line_id: str,
        *,
        qty: int | None = None,
        source_text: str | None = None,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        return self.draft_lifecycle.patch_order_line(
            draft_id,
            line_id,
            qty=qty,
            source_text=source_text,
            plate_order_ctx=plate_order_ctx,
        )

    def restore_order_lines(
        self,
        draft_id: str,
        *,
        index: int,
        lines: list[dict[str, Any]],
        replace_line_ids: list[str] | None = None,
    ) -> dict[str, Any]:
        return self.draft_lifecycle.restore_order_lines(
            draft_id,
            index=index,
            lines=lines,
            replace_line_ids=replace_line_ids,
        )

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
        return await self.product_draft_handler.create(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            owner_user_id=owner_user_id,
            plate_order_ctx=plate_order_ctx,
            product_type=product_type,
        )

    async def update_draft_piles(
        self,
        draft_id: str,
        *,
        mode: str,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.update(
            draft_id,
            product_type="piles",
            mode=mode,
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    async def apply_ai_piles_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.apply_ai(
            draft_id,
            product_type="piles",
            instruction=instruction,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    def update_draft_pile_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        return self.product_draft_handler.update_grades(
            draft_id,
            product_type="piles",
            concrete_grade=concrete_grade,
        )

    async def update_draft_marches(
        self,
        draft_id: str,
        *,
        mode: str,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.update(
            draft_id,
            product_type="marches",
            mode=mode,
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    async def apply_ai_marches_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.apply_ai(
            draft_id,
            product_type="marches",
            instruction=instruction,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    def update_draft_march_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        return self.product_draft_handler.update_grades(
            draft_id,
            product_type="marches",
            concrete_grade=concrete_grade,
        )

    async def update_draft_bridge_piles(
        self,
        draft_id: str,
        *,
        mode: str,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.update(
            draft_id,
            product_type="bridge_piles",
            mode=mode,
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    async def apply_ai_bridge_piles_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.apply_ai(
            draft_id,
            product_type="bridge_piles",
            instruction=instruction,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    def update_draft_bridge_pile_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        return self.product_draft_handler.update_grades(
            draft_id,
            product_type="bridge_piles",
            concrete_grade=concrete_grade,
        )

    async def update_draft_fbs(
        self,
        draft_id: str,
        *,
        mode: str,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.update(
            draft_id,
            product_type="fbs",
            mode=mode,
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    async def apply_ai_fbs_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.apply_ai(
            draft_id,
            product_type="fbs",
            instruction=instruction,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    def update_draft_fbs_grades(
        self,
        draft_id: str,
        *,
        concrete_grade: str,
    ) -> dict[str, Any]:
        return self.product_draft_handler.update_grades(
            draft_id,
            product_type="fbs",
            concrete_grade=concrete_grade,
        )

    async def update_draft_steps(
        self,
        draft_id: str,
        *,
        mode: str,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.update(
            draft_id,
            product_type="steps",
            mode=mode,
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    async def apply_ai_steps_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.apply_ai(
            draft_id,
            product_type="steps",
            instruction=instruction,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

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
        return await self.product_draft_handler.update(
            draft_id,
            product_type="plates",
            mode=mode,
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            plate_order_ctx=plate_order_ctx,
        )

    async def recognize_draft_page(
        self,
        draft_id: str,
        *,
        image_bytes: bytes,
        image_filename: str | None,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.recognize_page(
            draft_id,
            image_bytes=image_bytes,
            image_filename=image_filename,
        )

    async def apply_ai_plates_instruction(
        self,
        draft_id: str,
        *,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        return await self.product_draft_handler.apply_ai(
            draft_id,
            product_type="plates",
            instruction=instruction,
            image_bytes=image_bytes,
            image_filename=image_filename,
            plate_order_ctx=plate_order_ctx,
        )

    def resolve_wide_plates(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        return self.plate_resolve.resolve_wide_plates(
            draft_id,
            decisions,
            plate_order_ctx=plate_order_ctx,
        )

    def resolve_unpriced_plates(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        return self.plate_resolve.resolve_unpriced_plates(
            draft_id,
            decisions,
            plate_order_ctx=plate_order_ctx,
        )

    def resolve_invalid_widths(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        return self.plate_resolve.resolve_invalid_widths(
            draft_id,
            decisions,
            plate_order_ctx=plate_order_ctx,
        )

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
        pile_logistics_cost: float | None = None,
        pile_trip_overrides: dict[str, int] | None = None,
    ) -> dict[str, Any]:
        return self.draft_lifecycle.update_draft_meta(
            draft_id,
            manager_id=manager_id,
            client_name=client_name,
            discount_percent=discount_percent,
            conditions_mode=conditions_mode,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            logistics_cost=logistics_cost,
            pile_logistics_cost=pile_logistics_cost,
            pile_trip_overrides=pile_trip_overrides,
        )

    def calculate_draft(self, draft_id: str) -> dict[str, Any]:
        return self.draft_lifecycle.calculate_draft(draft_id)

    def get_draft_breakdown(
        self,
        draft_id: str,
        *,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        return self.draft_lifecycle.get_draft_breakdown(
            draft_id,
            plate_order_ctx=plate_order_ctx,
        )

    def get_draft_details(self, draft_id: str) -> dict[str, Any]:
        return self.draft_lifecycle.get_draft_details(draft_id)

    def hydrate_draft_from_saved_kp(
        self,
        kp_id: int,
        *,
        owner_user_id: int,
    ) -> dict[str, Any]:
        """Create a draft bound to an existing KP for append (status «в работе» only)."""
        return self.draft_lifecycle.hydrate_draft_from_saved_kp(
            kp_id, owner_user_id=owner_user_id
        )

    def _resolve_manager_for_hydrate(
        self,
        manager_name: str,
    ) -> tuple[int | None, str, str]:
        """Match managers.fio; keep a truthy manager_id when only the name is known."""
        return self.draft_lifecycle.resolve_manager_for_hydrate(manager_name)

    def generate_files(
        self,
        draft_id: str,
        file_types: Iterable[str] | None = None,
        *,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> list[dict[str, str]]:
        return self.draft_lifecycle.generate_files(
            draft_id,
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
        return self.draft_lifecycle.save_offer(
            draft_id,
            execution_terms=execution_terms,
            status=status,
            save_mode=save_mode,
        )

    def save_draft(self, draft_id: str, *, mode: str, execution_terms_input: str = "") -> dict[str, Any]:
        return self.draft_lifecycle.save_draft(
            draft_id, mode=mode, execution_terms_input=execution_terms_input
        )

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
        return self.draft_lifecycle.normalize_saved_offer(item)

    def _build_result_card(
        self,
        *,
        offer_identity: dict[str, str],
        metadata: dict[str, Any],
        saved_offer: dict[str, Any],
        totals: dict[str, Any],
    ) -> dict[str, Any]:
        return self.draft_lifecycle.build_result_card(
            offer_identity=offer_identity,
            metadata=metadata,
            saved_offer=saved_offer,
            totals=totals,
        )

    def _merge_plate_texts(self, current_text: str, next_text: str) -> str:
        parts = [current_text.strip(), next_text.strip()]
        return "\n".join(part for part in parts if part)

    def generate_and_persist_preview(
        self,
        *,
        text: str,
        owner_user_id: int,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        preview = self.commercial_service.generate_preview(
            text=text,
            plate_order_ctx=plate_order_ctx,
        )
        parse_result = preview.parse_result
        draft_id = self.draft_store.save_preview(
            order=parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=preview.order_data,
            metadata={
                "owner_user_id": owner_user_id,
                "normalized_text": parse_result.normalized_text,
                "warnings": parse_result.warnings,
                "unparsed_lines": parse_result.unparsed_lines,
                "normalized_lines": parse_result.normalized_lines,
                "wide_plate_lines": parse_result.wide_plate_lines,
                "diagnostics": parse_result.diagnostics,
                "breakdown_tables": preview.breakdown_tables,
                "price_rows_count": len(preview.price_rows),
                "breakdown_tables_count": len(preview.breakdown_tables),
                "total_sum": preview.total_sum,
            },
        )
        return {
            "draft_id": draft_id,
            "order": parse_result.order.to_dict(),
            "unparsed_lines": parse_result.unparsed_lines,
            "warnings": parse_result.warnings,
            "optimization": {
                "total_plates": preview.optimization_context.total_plates,
                "total_cost": preview.optimization_context.total_cost,
                "status": preview.optimization_context.optimization_status,
                "success": preview.optimization_context.optimization_success,
                "error_code": preview.optimization_context.optimization_error_code,
                "error_message": preview.optimization_context.optimization_error_message,
            },
            "order_data": preview.order_data,
            "price_rows_count": len(preview.price_rows),
            "breakdown_tables_count": len(preview.breakdown_tables),
            "total_sum": preview.total_sum,
        }

    def resolve_downloadable_file(self, draft_id: str, filename: str) -> Path:
        safe_name = Path(filename).name
        if safe_name != filename:
            raise ValueError("Некорректное имя файла.")
        if safe_name not in self.draft_store.generated_files_filenames(draft_id):
            raise FileNotFoundError("Файл не найден.")
        target_file = self._resolve_generated_file(safe_name).resolve()
        outputs_dir = Path(self.settings.outputs_dir).resolve()
        if target_file.parent != outputs_dir or not target_file.exists():
            raise FileNotFoundError("Файл не найден.")
        return target_file

    def get_or_generate_file(self, safe_filename: str) -> Path:
        """Path under configured ``outputs_dir`` for a generated file basename (no subpaths)."""
        return self.export_service.get_or_generate_file(safe_filename)

    def _resolve_generated_file(self, filename: str) -> Path:
        return self.export_service.resolve_generated_file(filename)
