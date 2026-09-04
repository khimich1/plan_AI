"""Draft calculate / meta / details / hydrate / export / save / append-cycle.

No workflow import: this module must stay free of CommercialWorkflowService.
Host is duck-typed (``CommercialWorkflowService`` instance).
"""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Any, Iterable

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from app.schemas.commercial import WizardStepId
from app.services.commercial_order_identity import APPEND_PRODUCT_TYPES
from app.services.product_draft_config import SPECS
from core.kp_order_data import order_data_from_kp_info
from core.pile_trip_pricing import coerce_pile_trip_overrides
from core.plate_order_context import PlateOrderContext

_PRODUCT_TYPE_TO_WIZARD_STEP = {key: spec.wizard_step for key, spec in SPECS.items()}
logger = logging.getLogger(__name__)


def recalc_promise_after_edit(db_path: str, kp_id: int) -> None:
    """Recalc active promise after constructor save. Occupancy errors are logged."""
    from app.services.promise_service import PromiseService
    from core.production.promise_buckets import OccupancyUnavailableError

    try:
        PromiseService(db_path=db_path).recalc_on_composition_change(int(kp_id))
    except OccupancyUnavailableError:
        logger.exception("promise recalc after edit: occupancy unavailable kp_id=%s", kp_id)
    except Exception:
        logger.exception("promise recalc after edit failed kp_id=%s", kp_id)


def _preview_unparsed_lines(preview: Any) -> list[str]:
    if hasattr(preview, "unparsed_lines"):
        return [str(item) for item in list(preview.unparsed_lines or []) if str(item).strip()]
    parse_result = getattr(preview, "parse_result", None)
    if parse_result is None:
        return []
    return [
        str(item)
        for item in list(getattr(parse_result, "unparsed_lines", None) or [])
        if str(item).strip()
    ]


def _scrub_ids_from_batches(
    batches: list[Any],
    remove_ids: set[str],
) -> list[dict[str, Any]]:
    next_batches: list[dict[str, Any]] = []
    for batch in batches:
        if not isinstance(batch, dict):
            continue
        next_batch = dict(batch)
        next_batch["line_ids"] = [
            str(lid)
            for lid in list(batch.get("line_ids") or [])
            if str(lid).strip() and str(lid).strip() not in remove_ids
        ]
        if next_batch["line_ids"]:
            next_batches.append(next_batch)
    return next_batches


def _replace_id_in_batches(
    batches: list[Any],
    old_id: str,
    new_ids: list[str],
) -> list[dict[str, Any]]:
    next_batches: list[dict[str, Any]] = []
    for batch in batches:
        if not isinstance(batch, dict):
            continue
        next_batch = dict(batch)
        ids = [str(lid) for lid in list(batch.get("line_ids") or [])]
        if old_id not in ids:
            next_batches.append(next_batch)
            continue
        idx = ids.index(old_id)
        next_batch["line_ids"] = ids[:idx] + new_ids + ids[idx + 1 :]
        if next_batch["line_ids"]:
            next_batches.append(next_batch)
    return next_batches


def _rebuild_append_batches(
    batches: list[Any],
    order_data: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Keep batch_id / product_type; set line_ids from current order_data order."""
    rebuilt: list[dict[str, Any]] = []
    seen: set[str] = set()
    for batch in batches:
        if not isinstance(batch, dict):
            continue
        batch_id = str(batch.get("batch_id") or "").strip()
        if not batch_id or batch_id in seen:
            continue
        seen.add(batch_id)
        line_ids = [
            str(line.get("line_id") or "").strip()
            for line in order_data
            if str(line.get("append_batch_id") or "").strip() == batch_id
            and str(line.get("line_id") or "").strip()
        ]
        if not line_ids:
            continue
        next_batch = dict(batch)
        next_batch["line_ids"] = line_ids
        rebuilt.append(next_batch)
    for line in order_data:
        batch_id = str(line.get("append_batch_id") or "").strip()
        if not batch_id or batch_id in seen:
            continue
        seen.add(batch_id)
        line_ids = [
            str(item.get("line_id") or "").strip()
            for item in order_data
            if str(item.get("append_batch_id") or "").strip() == batch_id
            and str(item.get("line_id") or "").strip()
        ]
        rebuilt.append(
            {
                "batch_id": batch_id,
                "product_type": str(line.get("product_type") or "plates").strip() or "plates",
                "line_ids": line_ids,
            }
        )
    return rebuilt


def _invalidate_breakdown(metadata: dict[str, Any]) -> None:
    """Drop cached breakdown so GET/export cannot serve pre-edit tables."""
    metadata["breakdown_tables"] = []
    metadata["breakdown_tables_count"] = 0
    files = list(metadata.get("generated_files") or [])
    metadata["generated_files"] = [
        item
        for item in files
        if not (isinstance(item, dict) and str(item.get("kind") or "").strip() == "breakdown")
    ]


def _plate_lines_text_from_order(order_data: list[dict[str, Any]]) -> str:
    """Rebuild plate ingest text from current order_data for breakdown refresh."""
    lines: list[str] = []
    for item in order_data:
        if not isinstance(item, dict):
            continue
        raw_type = str(item.get("product_type") or "").strip().lower()
        if raw_type and raw_type != "plates":
            continue
        source = str(item.get("source_text") or "").strip()
        if source:
            lines.append(source)
            continue
        name = str(item.get("name") or item.get("mark") or "").strip()
        qty = item.get("qty")
        if not name:
            continue
        if qty is None or qty == "":
            lines.append(name)
        else:
            lines.append(f"{name} {qty}")
    return "\n".join(lines)


class CommercialDraftLifecycle:
    def __init__(self, workflow: Any) -> None:
        self._wf = workflow

    def persist_order_and_metadata(
        self,
        draft_id: str,
        *,
        payload: dict[str, Any],
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any],
    ) -> None:
        self._wf.draft_store.replace_preview(
            draft_id,
            order=payload["order"],
            optimization_context=payload["optimization_context"],
            order_data=order_data,
            metadata=metadata,
        )

    def start_append_cycle(self, draft_id: str, *, product_type: str) -> dict[str, Any]:
        """Switch cycle product_type, clear cycle input, keep header + prior order_data."""
        normalized = (product_type or "").strip().lower()
        if normalized not in APPEND_PRODUCT_TYPES:
            raise ValueError("Некорректный тип продукта.")

        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata") or {})
        order_data = list(payload.get("order_data") or [])

        sealed_order, sealed_batches = self._wf.order_identity.seal_unbatched_lines(
            order_data, metadata
        )
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
        metadata["invalid_width_lines"] = []
        metadata["invalid_widths_resolved"] = True
        metadata["unparsed_lines"] = []
        metadata["warnings"] = []
        metadata["last_source_filename"] = ""
        metadata["source_type"] = None

        wizard_step = _PRODUCT_TYPE_TO_WIZARD_STEP.get(normalized, WizardStepId.plates)
        metadata["current_step"] = wizard_step.value

        self.persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=sealed_order,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def undo_last_append_batch(self, draft_id: str) -> dict[str, Any]:
        """Remove the last append_batches entry and its lines from order_data."""
        payload = self._wf._load_draft_or_raise(draft_id)
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
        _invalidate_breakdown(metadata)
        self.persist_order_and_metadata(
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

        payload = self._wf._load_draft_or_raise(draft_id)
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
        _invalidate_breakdown(metadata)
        self.persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=next_order,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def patch_order_line(
        self,
        draft_id: str,
        line_id: str,
        *,
        qty: int | None = None,
        source_text: str | None = None,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        """Update qty in place, or replace the line with fragment preview of source_text."""
        target_id = (line_id or "").strip()
        if not target_id:
            raise ValueError("Не указан идентификатор строки.")
        text = (source_text or "").strip() if source_text is not None else None
        if qty is None and not text:
            raise ValueError("Укажите количество или текст строки.")
        if qty is not None and int(qty) < 1:
            raise ValueError("Количество должно быть больше нуля.")

        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata") or {})
        order_data = [
            dict(line)
            for line in list(payload.get("order_data") or [])
            if isinstance(line, dict)
        ]
        index = next(
            (
                i
                for i, line in enumerate(order_data)
                if str(line.get("line_id") or "").strip() == target_id
            ),
            None,
        )
        if index is None:
            raise FileNotFoundError("Строка не найдена.")

        if text:
            order_data, metadata = self._replace_line_from_source_text(
                order_data,
                metadata,
                index=index,
                source_text=text,
                plate_order_ctx=plate_order_ctx,
            )
        else:
            self._apply_qty_to_line(order_data[index], int(qty))

        _invalidate_breakdown(metadata)
        self.persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=order_data,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def restore_order_lines(
        self,
        draft_id: str,
        *,
        index: int,
        lines: list[dict[str, Any]],
        replace_line_ids: list[str] | None = None,
    ) -> dict[str, Any]:
        """Remove replace_line_ids, then splice snapshot lines at index."""
        snapshots = [dict(line) for line in list(lines or []) if isinstance(line, dict)]
        if not snapshots:
            raise ValueError("Нет строк для восстановления.")
        insert_at = max(0, int(index))

        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata") or {})
        order_data = [
            dict(line)
            for line in list(payload.get("order_data") or [])
            if isinstance(line, dict)
        ]
        remove_ids = {
            str(lid).strip()
            for lid in list(replace_line_ids or [])
            if str(lid).strip()
        }
        if remove_ids:
            order_data = [
                line
                for line in order_data
                if str(line.get("line_id") or "").strip() not in remove_ids
            ]
            metadata["append_batches"] = _scrub_ids_from_batches(
                list(metadata.get("append_batches") or []),
                remove_ids,
            )

        insert_at = min(insert_at, len(order_data))
        order_data = order_data[:insert_at] + snapshots + order_data[insert_at:]
        metadata["append_batches"] = _rebuild_append_batches(
            list(metadata.get("append_batches") or []),
            order_data,
        )
        _invalidate_breakdown(metadata)
        self.persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=order_data,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def _replace_line_from_source_text(
        self,
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any],
        *,
        index: int,
        source_text: str,
        plate_order_ctx: PlateOrderContext | None,
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        from app.services.product_draft_config import get_spec

        old_line = order_data[index]
        old_id = str(old_line.get("line_id") or "").strip()
        product_type = (
            str(old_line.get("product_type") or metadata.get("product_type") or "plates")
            .strip()
            .lower()
            or "plates"
        )
        spec = get_spec(product_type)
        preview = spec.generate_preview(
            self._wf,
            source_text,
            plate_order_ctx=plate_order_ctx,
        )
        if _preview_unparsed_lines(preview):
            raise ValueError("Не удалось распознать строку.")
        preview_rows = [
            dict(row)
            for row in list(getattr(preview, "order_data", None) or [])
            if isinstance(row, dict)
        ]
        stamped = self._wf.order_identity.stamp_order_data(
            preview_rows,
            product_type=product_type,
        )
        batch_id = str(old_line.get("append_batch_id") or "").strip()
        if batch_id:
            for row in stamped:
                row["append_batch_id"] = batch_id
        next_order = order_data[:index] + stamped + order_data[index + 1 :]
        new_ids = [
            str(row.get("line_id") or "").strip()
            for row in stamped
            if str(row.get("line_id") or "").strip()
        ]
        metadata["append_batches"] = _replace_id_in_batches(
            list(metadata.get("append_batches") or []),
            old_id,
            new_ids,
        )
        return next_order, metadata

    @staticmethod
    def _apply_qty_to_line(line: dict[str, Any], qty: int) -> None:
        old_qty = float(line.get("qty") or 0)
        line["qty"] = qty
        unit = line.get("unit_price")
        if unit is not None:
            line["line_total"] = float(unit) * qty
        if line.get("length_m") is not None or line.get("width_m") is not None:
            from core.kp_plate_weight import resolve_kp_line_weight_kg

            _, total_kg = resolve_kp_line_weight_kg(line)
            line["weight"] = total_kg
        elif old_qty > 0 and line.get("weight") is not None:
            line["weight"] = float(line["weight"]) / old_qty * qty

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
        payload_before = self._wf._load_draft_or_raise(draft_id)
        prev_step = self._wf._normalize_stored_step(dict(payload_before.get("metadata") or {}))

        updates: dict[str, Any] = {}
        if manager_id is not None:
            manager = self._wf.manager_repository.get_manager(manager_id)
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
        if pile_logistics_cost is not None:
            if pile_logistics_cost < 0:
                raise ValueError("Стоимость рейса свай не может быть отрицательной.")
            updates["pile_logistics_cost"] = float(pile_logistics_cost)
        if pile_trip_overrides is not None:
            from core.pile_trip_pricing import coerce_pile_trip_overrides

            updates["pile_trip_overrides"] = coerce_pile_trip_overrides(pile_trip_overrides)
        if updates:
            self._wf.draft_store.update_metadata(draft_id, **updates)

        payload_after = self._wf._load_draft_or_raise(draft_id)
        md = dict(payload_after.get("metadata") or {})

        financial_keys = {
            "discount_percent",
            "logistics_cost",
            "pile_logistics_cost",
            "pile_trip_overrides",
        }
        if updates:
            if prev_step == WizardStepId.result and set(updates.keys()).issubset(financial_keys):
                self._wf._persist_wizard_step(draft_id, WizardStepId.result)
            else:
                self._wf._persist_wizard_step(draft_id, WizardStepId.client)

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
            stamped = self._wf.order_identity.stamp_order_data(
                order_data,
                product_type=product_type,
                previous_order_data=order_data,
            )
            payload = self._wf._load_draft_or_raise(draft_id)
            self._wf.draft_store.replace_preview(
                draft_id,
                order=payload["order"],
                optimization_context=payload["optimization_context"],
                order_data=stamped,
                metadata=dict(payload.get("metadata") or {}),
            )
            order_data = stamped
        self._wf.calculation_service.enforce_calculate_prerequisites(
            order_data=order_data,
            metadata=dict(metadata),
        )
        payload = self._wf._load_draft_or_raise(draft_id)
        meta = dict(payload.get("metadata") or {})
        sealed_order, sealed_batches = self._wf.order_identity.seal_unbatched_lines(
            list(payload.get("order_data") or order_data),
            meta,
        )
        meta["append_batches"] = sealed_batches
        meta["current_step"] = WizardStepId.result.value
        self.persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=sealed_order,
            metadata=meta,
        )
        return self.get_draft_details(draft_id)

    def get_draft_breakdown(
        self,
        draft_id: str,
        *,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> dict[str, Any]:
        self.refresh_breakdown_if_needed(draft_id, plate_order_ctx=plate_order_ctx)
        payload = self._wf._load_draft_or_raise(draft_id)
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

    def refresh_breakdown_if_needed(
        self,
        draft_id: str,
        *,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> None:
        """When tables are empty after invalidate, rebuild from current plate lines."""
        if plate_order_ctx is None:
            return
        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata") or {})
        if list(metadata.get("breakdown_tables") or []):
            return
        order_data = [
            dict(line)
            for line in list(payload.get("order_data") or [])
            if isinstance(line, dict)
        ]
        plate_text = _plate_lines_text_from_order(order_data)
        if not plate_text.strip():
            return
        commercial_service = getattr(self._wf, "commercial_service", None)
        if commercial_service is None:
            return
        try:
            preview = commercial_service.generate_preview(
                text=plate_text,
                plate_order_ctx=plate_order_ctx,
            )
        except Exception:
            return
        tables = list(getattr(preview, "breakdown_tables", None) or [])
        if not tables:
            return
        metadata["breakdown_tables"] = tables
        metadata["breakdown_tables_count"] = len(tables)
        self.persist_order_and_metadata(
            draft_id,
            payload=payload,
            order_data=order_data,
            metadata=metadata,
        )

    def get_draft_details(self, draft_id: str) -> dict[str, Any]:
        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        totals = self._wf.calculation_service.compute_totals_from_metadata(
            payload["order_data"],
            metadata,
        )
        public_metadata = {
            key: value
            for key, value in metadata.items()
            if key not in ("breakdown_tables", "owner_user_id", "schema_file")
        }
        wizard_state = self._wf.build_wizard_state(payload)
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
            "files": self._wf.export_service.collect_draft_files(metadata, draft_id),
            "saved_offer": self.normalize_saved_offer(metadata.get("saved_offer")),
            "totals": totals,
            "offer_identity": self._wf.export_service.build_offer_identity_payload(draft_id),
        }

    def hydrate_draft_from_saved_kp(
        self,
        kp_id: int,
        *,
        owner_user_id: int,
    ) -> dict[str, Any]:
        """Create a draft bound to an existing KP for append (status «в архиве» only)."""
        kp_raw = self._wf.kp_repository.get_offer(kp_id)
        if not kp_raw:
            raise ValueError(f"КП №{kp_id} не найдено")

        status = str(kp_raw.get("status") or "").strip()
        if status != "в архиве":
            raise ValueError("Дополнить КП можно только в статусе «в архиве».")

        order_data = [
            dict(line)
            for line in order_data_from_kp_info(kp_raw)
            if isinstance(line, dict)
        ]
        cycle_type = "plates"
        for line in reversed(order_data):
            pt = str(line.get("product_type") or "").strip().lower()
            if pt in APPEND_PRODUCT_TYPES:
                cycle_type = pt
                break

        manager_name = str(kp_raw.get("manager_name") or "").strip()
        manager_id, manager_phone, manager_email = self.resolve_manager_for_hydrate(
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
            "pile_logistics_cost": float(kp_raw.get("pile_logistics_cost") or 0.0),
            "pile_trip_overrides": coerce_pile_trip_overrides(
                kp_raw.get("pile_trip_overrides_json")
            ),
            "delivery_conditions": delivery,
            "payment_conditions": payment,
            "conditions_mode": conditions_mode,
            "execution_terms": execution_terms,
            "wide_plates_resolved": True,
            "wide_plate_lines": [],
            "invalid_widths_resolved": True,
            "invalid_width_lines": [],
            "append_batches": [],
            "resume_kp_id": int(kp_id),
            "current_step": WizardStepId.result.value,
            "saved_offer": saved_offer,
        }
        draft_id = self._wf.draft_store.save_preview(
            order=PlateOrder(),
            optimization_context=OptimizationContext(order=PlateOrder()),
            order_data=order_data,
            metadata=metadata,
        )
        return self.get_draft_details(draft_id)

    def resolve_manager_for_hydrate(
        self,
        manager_name: str,
    ) -> tuple[int | None, str, str]:
        """Match managers.fio; keep a truthy manager_id when only the name is known."""
        name = (manager_name or "").strip()
        if not name:
            return None, "", ""
        needle = name.casefold()
        for manager in self._wf.manager_repository.list_managers():
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
        self.refresh_breakdown_if_needed(draft_id, plate_order_ctx=plate_order_ctx)
        payload = self._wf._load_draft_or_raise(draft_id)
        return self._wf.export_service.generate_files(
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
        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        files = self._wf.generate_files(draft_id, ("xlsx",))
        xlsx_file = next((item for item in files if item["kind"] == "xlsx"), None)
        xlsx_path = None
        if xlsx_file:
            resolved = self._wf.export_service.resolve_generated_file(xlsx_file["filename"])
            xlsx_path = str(resolved) if resolved.exists() else None

        raw_owner = metadata.get("owner_user_id")
        owner_user_id = int(raw_owner) if raw_owner is not None else None
        customer_name = str(metadata.get("client_name", "") or "Клиент")
        manager_name = str(metadata.get("manager_name", "") or "")
        discount_percent = float(metadata.get("discount_percent", 0.0) or 0.0)
        logistics_cost = float(metadata.get("logistics_cost", 0.0) or 0.0)
        pile_logistics_cost = float(metadata.get("pile_logistics_cost", 0.0) or 0.0)
        pile_trip_overrides = coerce_pile_trip_overrides(
            metadata.get("pile_trip_overrides")
        )
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
                    "status": existing_saved.get("status") or "в архиве",
                }
        if existing_kp_id is not None:
            existing_status = str(existing_saved.get("status", "") or "").strip()
            if existing_status != "в архиве":
                raise ValueError(
                    "Дополнить КП можно только в статусе «в архиве»."
                )
            kp_id = self._wf.kp_repository.update_offer_from_order_data(
                int(existing_kp_id),
                order_data=order_data,
                customer_name=customer_name,
                manager_name=manager_name,
                discount_percent=discount_percent,
                logistics_cost=logistics_cost,
                pile_logistics_cost=pile_logistics_cost,
                pile_trip_overrides=pile_trip_overrides,
                delivery_conditions=delivery_conditions,
                payment_conditions=payment_conditions,
                execution_terms=execution_terms,
                xlsx_path=xlsx_path,
                product_type=product_type,
            )
            recalc_promise_after_edit(str(self._wf.kp_repository.db_path), int(kp_id))
            # Keep archived status on update; do not flip to default «в работе».
            persist_status = existing_status
        else:
            kp_id = self._wf.kp_repository.save_offer(
                creation_date=datetime.now().strftime("%d.%m.%Y"),
                customer_name=customer_name,
                manager_name=manager_name,
                discount_percent=discount_percent,
                logistics_cost=logistics_cost,
                pile_logistics_cost=pile_logistics_cost,
                pile_trip_overrides=pile_trip_overrides,
                delivery_conditions=delivery_conditions,
                payment_conditions=payment_conditions,
                execution_terms=execution_terms,
                status=status,
                order_data=order_data,
                xlsx_path=xlsx_path,
                owner_user_id=owner_user_id,
                product_type=product_type,
            )
            persist_status = status
        saved_offer = {
            "kp_id": kp_id,
            "status": persist_status,
            "mode": save_mode,
            "execution_terms": execution_terms,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }
        self._wf.draft_store.update_metadata(
            draft_id,
            saved_offer=saved_offer,
            current_save_mode=save_mode,
            execution_terms=execution_terms,
        )
        totals = self._wf.calculation_service.compute_totals_from_metadata(
            payload["order_data"],
            metadata,
        )
        offer_identity = self._wf.export_service.build_offer_identity_payload(draft_id)
        return {
            "saved_offer": saved_offer,
            "totals": totals,
            "offer_identity": offer_identity,
            "result_card": self.build_result_card(
                offer_identity=offer_identity,
                metadata=metadata,
                saved_offer=saved_offer,
                totals=totals,
            ),
        }

    def save_draft(self, draft_id: str, *, mode: str, execution_terms_input: str = "") -> dict[str, Any]:
        normalized_mode = mode.strip().lower()
        if normalized_mode == "database":
            execution_terms = self._wf.execution_terms_service.normalize(execution_terms_input)
            return self.save_offer(
                draft_id,
                execution_terms=execution_terms,
                status="в работе",
                save_mode="database",
            )
        if normalized_mode == "archive":
            raw = (execution_terms_input or "").strip()
            execution_terms = self._wf.execution_terms_service.normalize(raw) if raw else ""
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
            self._wf.draft_store.update_metadata(draft_id, saved_offer=skipped_offer, current_save_mode="skip")
            return {
                "saved_offer": skipped_offer,
                "totals": details["totals"],
                "offer_identity": details["offer_identity"],
                "result_card": self.build_result_card(
                    offer_identity=details["offer_identity"],
                    metadata=details["metadata"],
                    saved_offer=skipped_offer,
                    totals=details["totals"],
                ),
            }
        raise ValueError("Некорректный режим сохранения.")

    @staticmethod
    def normalize_saved_offer(item: dict[str, Any] | None) -> dict[str, Any] | None:
        if not item:
            return None
        return {
            "kp_id": item.get("kp_id"),
            "status": str(item.get("status", "") or ""),
            "mode": str(item.get("mode", "database") or "database"),
            "execution_terms": str(item.get("execution_terms", "") or ""),
            "saved_at": str(item.get("saved_at", "") or ""),
        }

    @staticmethod
    def build_result_card(
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
