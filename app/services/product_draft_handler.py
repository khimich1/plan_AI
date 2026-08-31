"""Config-driven create / update / AI / grades for commercial product drafts."""

from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from app.services.commercial_draft_service import _safe_ocr_temp_suffix
from app.services.product_draft_config import ProductDraftSpec, get_spec
from core.plate_order_context import PlateOrderContext


class ProductDraftHandler:
    def __init__(self, workflow: Any) -> None:
        self._wf = workflow
        self._id = workflow.order_identity

    async def create(
        self,
        *,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        owner_user_id: int,
        plate_order_ctx: PlateOrderContext | None = None,
        product_type: str = "plates",
    ) -> dict[str, Any]:
        spec = get_spec((product_type or "plates").strip().lower())
        if spec.allow_empty_create and not (text or "").strip() and not image_bytes:
            return self._create_empty(spec, owner_user_id)

        source_text, source_metadata = await self._wf.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type=spec.product_type,
        )
        preview = spec.generate_preview(
            self._wf,
            source_text["input_text"],
            plate_order_ctx=plate_order_ctx,
        )
        batches = [source_text["batch"]]
        metadata = self._build_metadata(
            spec,
            preview=preview,
            base_metadata={"product_type": spec.product_type} if not spec.use_preview_order else {},
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=source_text["input_text"],
            last_source_filename=source_text["filename"],
            batches=batches,
            source_metadata=source_metadata,
            owner_user_id=owner_user_id,
            preview_for_wide=preview,
        )
        order_data = self._id.stamp_order_data(preview.order_data, product_type=spec.product_type)
        draft_id = self._save_preview(spec, preview, order_data, metadata)
        self._persist_wizard_after_write(spec, draft_id)
        return self._wf.get_draft_details(draft_id)

    async def update(
        self,
        draft_id: str,
        *,
        product_type: str,
        mode: str,
        text: str | None,
        image_bytes: bytes | None,
        image_filename: str | None,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        spec = get_spec(product_type)
        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        self._guard_update(spec, metadata)

        source_text, source_metadata = await self._wf.draft_service.resolve_source_input(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_filename,
            product_type=spec.product_type,
        )
        current_text = str(metadata.get("input_text", "") or "")
        normalized_mode = mode.strip().lower()
        if normalized_mode not in {"append", "replace"}:
            raise ValueError(spec.invalid_mode_error)
        merged_cycle_text = bool(normalized_mode == "append" and current_text.strip())
        if merged_cycle_text:
            next_text = self._wf._merge_plate_texts(current_text, source_text["input_text"])
            batches = list(metadata.get(spec.batches_key) or [])
            batches.append(source_text["batch"])
        else:
            next_text = source_text["input_text"]
            batches = [source_text["batch"]]

        preview = spec.generate_preview(self._wf, next_text, plate_order_ctx=plate_order_ctx)
        next_metadata = self._build_metadata(
            spec,
            preview=preview,
            base_metadata=metadata,
            source_type=source_text["source_type"],
            original_text=source_text["original_text"],
            ocr_text=source_text["ocr_text"],
            input_text=next_text,
            last_source_filename=source_text["filename"],
            batches=batches,
            source_metadata=source_metadata,
            preview_for_wide=preview,
        )
        order_data = self._compose_stamped(
            spec,
            previous_order_data=list(payload.get("order_data") or []),
            preview_order_data=preview.order_data,
            mode=normalized_mode,
            merged_cycle_text=merged_cycle_text,
        )
        self._replace_preview(spec, draft_id, preview, order_data, next_metadata)
        self._persist_wizard_after_write(spec, draft_id)
        return self._wf.get_draft_details(draft_id)

    async def apply_ai(
        self,
        draft_id: str,
        *,
        product_type: str,
        instruction: str,
        image_bytes: bytes | None,
        image_filename: str | None,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        spec = get_spec(product_type)
        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if spec.type_mismatch_ai is not None:
            if str(metadata.get("product_type", "plates")).lower() != spec.product_type:
                raise ValueError(spec.type_mismatch_ai)

        instruction_value = (instruction or "").strip()
        if len(instruction_value) < 3:
            raise ValueError("Инструкция для ИИ должна содержать минимум 3 символа.")

        current_text = str(metadata.get("input_text", "") or "")
        result = await self._run_ai(spec, current_text, instruction_value, image_bytes, image_filename)
        recognized_text, source_metadata, ai_batch = self._ai_payload(
            spec,
            result,
            instruction_value=instruction_value,
            image_filename=image_filename,
        )
        preview = spec.generate_preview(self._wf, recognized_text, plate_order_ctx=plate_order_ctx)
        next_metadata = self._build_metadata(
            spec,
            preview=preview,
            base_metadata=metadata,
            source_type="ai",
            original_text=ai_batch.get("original_text", instruction_value)
            if spec.ai_payload_style == "rich"
            else str(metadata.get("original_text", "") or ""),
            ocr_text=recognized_text,
            input_text=recognized_text,
            last_source_filename=image_filename or "",
            batches=[ai_batch],
            source_metadata=source_metadata,
            preview_for_wide=preview,
        )
        order_data = self._id.stamp_order_data(preview.order_data, product_type=spec.product_type)
        self._replace_preview(spec, draft_id, preview, order_data, next_metadata)
        self._persist_wizard_after_write(spec, draft_id)
        return self._wf.get_draft_details(draft_id)

    def update_grades(
        self,
        draft_id: str,
        *,
        product_type: str,
        concrete_grade: str,
    ) -> dict[str, Any]:
        spec = get_spec(product_type)
        if not spec.has_grades or spec.type_mismatch_grades is None or spec.grades_empty_error is None:
            raise ValueError("Некорректный тип продукта.")

        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        if str(metadata.get("product_type", "plates")).lower() != spec.product_type:
            raise ValueError(spec.type_mismatch_grades)

        grade = (concrete_grade or "").strip()
        if not grade:
            raise ValueError("Укажите класс бетона.")

        previous_order_data = list(payload.get("order_data") or [])
        cycle_items = self._id.current_cycle_lines(previous_order_data, product_type=spec.product_type)
        if not cycle_items:
            raise ValueError(spec.grades_empty_error)

        lines, skipped = self._grade_lines(spec, cycle_items, grade)
        if not lines:
            raise ValueError(spec.grades_empty_error)

        next_text = "\n".join(lines)
        preview = spec.generate_preview(self._wf, next_text, plate_order_ctx=None)
        next_metadata = self._build_metadata(
            spec,
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            batches=list(metadata.get(spec.batches_key) or []),
            source_metadata={},
            preview_for_wide=preview,
        )
        next_metadata["default_concrete_grade"] = grade
        self._apply_grade_skip_warnings(spec, next_metadata, grade, skipped)

        stamp_previous = self._id.stamp_previous_for_product_update(
            cycle_items,
            mode="append",
            merged_cycle_text=True,
        )
        new_type_lines = self._id.stamp_order_data(
            preview.order_data,
            product_type=spec.product_type,
            previous_order_data=stamp_previous,
        )
        order_data = self._id.compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_type_lines,
            product_type=spec.product_type,
            mode="append",
            merged_cycle_text=True,
        )
        self._replace_preview(spec, draft_id, preview, order_data, next_metadata)
        self._persist_wizard_after_write(spec, draft_id)
        return self._wf.get_draft_details(draft_id)

    def _create_empty(self, spec: ProductDraftSpec, owner_user_id: int) -> dict[str, Any]:
        metadata: dict[str, Any] = {
            "product_type": spec.product_type,
            "owner_user_id": owner_user_id,
            "current_step": spec.wizard_step.value,
            "wide_plates_resolved": True,
        }
        if spec.empty_create_grade is not None:
            metadata["default_concrete_grade"] = spec.empty_create_grade
        draft_id = self._wf.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=[],
            metadata=metadata,
        )
        return self._wf.get_draft_details(draft_id)

    @staticmethod
    def _guard_update(spec: ProductDraftSpec, metadata: dict[str, Any]) -> None:
        current = str(metadata.get("product_type", "plates")).lower()
        if spec.update_reject_map is not None:
            message = spec.update_reject_map.get(current)
            if message:
                raise ValueError(message)
            return
        if current != spec.product_type:
            raise ValueError(spec.type_mismatch_update)

    def _build_metadata(
        self,
        spec: ProductDraftSpec,
        *,
        preview: Any,
        base_metadata: dict[str, Any],
        source_type: str,
        original_text: str,
        ocr_text: str,
        input_text: str,
        last_source_filename: str,
        batches: list[dict[str, Any]],
        source_metadata: dict[str, Any],
        owner_user_id: int | None = None,
        preview_for_wide: Any = None,
    ) -> dict[str, Any]:
        kwargs: dict[str, Any] = {
            "preview": preview,
            "base_metadata": base_metadata,
            "source_type": source_type,
            "original_text": original_text,
            "ocr_text": ocr_text,
            "input_text": input_text,
            "last_source_filename": last_source_filename,
            "batches": batches,
            "source_metadata": source_metadata,
            "owner_user_id": owner_user_id,
        }
        if spec.needs_plate_ctx:
            wide_src = preview_for_wide if preview_for_wide is not None else preview
            kwargs["wide_plates_resolved"] = not bool(wide_src.parse_result.wide_plate_lines)
        return spec.build_metadata(self._wf.draft_service, **kwargs)

    def _compose_stamped(
        self,
        spec: ProductDraftSpec,
        *,
        previous_order_data: list[Any],
        preview_order_data: list[Any],
        mode: str,
        merged_cycle_text: bool,
    ) -> list[dict[str, Any]]:
        _, same_previous = self._id.partition_order_by_product_type(
            previous_order_data,
            product_type=spec.product_type,
        )
        stamp_previous = self._id.stamp_previous_for_product_update(
            same_previous,
            mode=mode,
            merged_cycle_text=merged_cycle_text,
        )
        new_type_lines = self._id.stamp_order_data(
            preview_order_data,
            product_type=spec.product_type,
            previous_order_data=stamp_previous,
        )
        return self._id.compose_order_data_for_product_update(
            previous_order_data=previous_order_data,
            new_type_lines=new_type_lines,
            product_type=spec.product_type,
            mode=mode,
            merged_cycle_text=merged_cycle_text,
        )

    def _save_preview(
        self,
        spec: ProductDraftSpec,
        preview: Any,
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any],
    ) -> str:
        order, opt = self._preview_order(spec, preview)
        return self._wf.draft_store.save_preview(
            order=order,
            optimization_context=opt,
            order_data=order_data,
            metadata=metadata,
        )

    def _replace_preview(
        self,
        spec: ProductDraftSpec,
        draft_id: str,
        preview: Any,
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any],
    ) -> None:
        order, opt = self._preview_order(spec, preview)
        self._wf.draft_store.replace_preview(
            draft_id,
            order=order,
            optimization_context=opt,
            order_data=order_data,
            metadata=metadata,
        )

    @staticmethod
    def _preview_order(spec: ProductDraftSpec, preview: Any) -> tuple[Any, Any]:
        if spec.use_preview_order:
            return preview.parse_result.order, preview.optimization_context
        empty = PlateOrder()
        return empty, OptimizationContext(order=PlateOrder())

    def _persist_wizard_after_write(self, spec: ProductDraftSpec, draft_id: str) -> None:
        if spec.use_preview_order:
            payload_snap = self._wf._load_draft_or_raise(draft_id)
            step = self._wf._wizard_step_after_plate_snapshot(
                dict(payload_snap.get("metadata", {})),
                payload_snap["order_data"],
            )
            self._wf._persist_wizard_step(draft_id, step)
            return
        self._wf._persist_wizard_step(draft_id, spec.wizard_step)

    async def _run_ai(
        self,
        spec: ProductDraftSpec,
        current_text: str,
        instruction_value: str,
        image_bytes: bytes | None,
        image_filename: str | None,
    ) -> Any:
        tmp_path: Path | None = None
        try:
            image_path: str | None = None
            if image_bytes:
                suffix = _safe_ocr_temp_suffix(image_filename)
                with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(image_bytes)
                    tmp_path = Path(tmp_file.name)
                image_path = str(tmp_path)
            return await spec.apply_ai(current_text, instruction_value, image_path)
        finally:
            if tmp_path is not None:
                if spec.ai_payload_style == "compact":
                    try:
                        tmp_path.unlink(missing_ok=True)
                    except OSError:
                        pass
                else:
                    tmp_path.unlink(missing_ok=True)

    def _ai_payload(
        self,
        spec: ProductDraftSpec,
        result: Any,
        *,
        instruction_value: str,
        image_filename: str | None,
    ) -> tuple[str, dict[str, Any], dict[str, Any]]:
        if spec.ai_payload_style == "compact":
            recognized_text = str(result.get("text") or "").strip()
            if not recognized_text:
                raise ValueError(spec.ai_empty_error)
            ai_batch = {
                "source_type": "ai",
                "filename": image_filename or "",
                "text": recognized_text,
            }
            source_metadata = {
                "ai_applied": True,
                "ocr_warnings": list(result.get("warnings") or []),
            }
            return recognized_text, source_metadata, ai_batch

        recognized_text = str((result or {}).get("text", "")).strip()
        if not recognized_text:
            raise ValueError(spec.ai_empty_error)
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
        return recognized_text, source_metadata, ai_batch

    @staticmethod
    def _grade_lines(
        spec: ProductDraftSpec,
        cycle_items: list[dict[str, Any]],
        grade: str,
    ) -> tuple[list[str], list[str]]:
        lines: list[str] = []
        skipped: list[str] = []
        if not spec.grades_skip_unavailable:
            for item in cycle_items:
                mark = str(item.get("mark") or item.get("name") or "").strip()
                qty = int(item.get("qty") or 0)
                if mark and qty > 0:
                    lines.append(f"{mark} {grade} {qty}")
            return lines, skipped

        list_grades = spec.list_available_grades
        for item in cycle_items:
            mark = str(item.get("mark") or item.get("name") or "").strip()
            qty = int(item.get("qty") or 0)
            if not mark or qty <= 0:
                continue
            available = list(item.get("available_grades") or [])
            if not available and list_grades is not None:
                available = list_grades(mark)
            if grade in available:
                lines.append(f"{mark} {grade} {qty}")
            else:
                current_grade = str(item.get("concrete_grade") or "B25").strip()
                lines.append(f"{mark} {current_grade} {qty}")
                skipped.append(mark)
        return lines, skipped

    @staticmethod
    def _apply_grade_skip_warnings(
        spec: ProductDraftSpec,
        next_metadata: dict[str, Any],
        grade: str,
        skipped: list[str],
    ) -> None:
        if not spec.grades_skip_unavailable:
            return
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
