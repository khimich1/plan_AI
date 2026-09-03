"""Unified wide / unpriced / invalid-width plate resolve for commercial drafts.

No workflow import: this module must stay free of CommercialWorkflowService.
Host is duck-typed (``CommercialWorkflowService`` instance).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable, Literal

from app.schemas.commercial import WizardStepId
from core.plate_order_context import PlateOrderContext


@dataclass(frozen=True)
class PlateResolveSpec:
    kind: Literal["wide", "unpriced", "invalid_width"]
    unresolved_error: str
    empty_after_error: str
    invalid_action_error: str
    decisions_meta_key: str
    extra_resolved_flag: str | None
    force_wide_resolved: bool
    coerce_blank_source_line: bool


WIDE_PLATE_RESOLVE = PlateResolveSpec(
    kind="wide",
    unresolved_error="Нужно выбрать действие для всех широких плит.",
    empty_after_error="После обработки широких плит список стал пустым.",
    invalid_action_error="Некорректное действие для обработки широкой плиты.",
    decisions_meta_key="wide_plate_decisions",
    extra_resolved_flag=None,
    force_wide_resolved=True,
    coerce_blank_source_line=False,
)
UNPRICED_PLATE_RESOLVE = PlateResolveSpec(
    kind="unpriced",
    unresolved_error="Нужно выбрать действие для всех позиций без цены.",
    empty_after_error="После обработки позиций без цены список стал пустым.",
    invalid_action_error="Некорректное действие для позиции без цены.",
    decisions_meta_key="unpriced_plate_decisions",
    extra_resolved_flag="unpriced_plates_resolved",
    force_wide_resolved=False,
    coerce_blank_source_line=True,
)
INVALID_WIDTH_RESOLVE = PlateResolveSpec(
    kind="invalid_width",
    unresolved_error="Нужно выбрать действие для всех позиций с нестандартной шириной.",
    empty_after_error="После обработки нестандартной ширины список стал пустым.",
    invalid_action_error="Некорректное действие для позиции с нестандартной шириной.",
    decisions_meta_key="invalid_width_decisions",
    extra_resolved_flag="invalid_widths_resolved",
    force_wide_resolved=False,
    coerce_blank_source_line=True,
)


class CommercialPlateResolve:
    def __init__(self, workflow: Any) -> None:
        self._wf = workflow

    def resolve_wide_plates(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        return self.resolve_plate_items(
            draft_id,
            decisions,
            plate_order_ctx=plate_order_ctx,
            spec=WIDE_PLATE_RESOLVE,
        )

    def resolve_unpriced_plates(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        return self.resolve_plate_items(
            draft_id,
            decisions,
            plate_order_ctx=plate_order_ctx,
            spec=UNPRICED_PLATE_RESOLVE,
        )

    def resolve_invalid_widths(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> dict[str, Any]:
        return self.resolve_plate_items(
            draft_id,
            decisions,
            plate_order_ctx=plate_order_ctx,
            spec=INVALID_WIDTH_RESOLVE,
        )

    def resolve_plate_items(
        self,
        draft_id: str,
        decisions: Iterable[dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
        spec: PlateResolveSpec,
    ) -> dict[str, Any]:
        payload = self._wf._load_draft_or_raise(draft_id)
        metadata = dict(payload.get("metadata", {}))
        current_text = str(metadata.get("input_text", "") or "")
        if not current_text:
            raise ValueError("Список плит отсутствует.")

        items = self._load_plate_resolve_items(spec, metadata)
        if not items:
            return self._wf.get_draft_details(draft_id)

        decisions_by_id, decisions_by_line = self._index_plate_resolve_decisions(
            decisions,
            coerce_blank_source_line=spec.coerce_blank_source_line,
        )
        resolved_by_line, resolved_decisions = self._bind_plate_resolve_items(
            spec,
            items,
            decisions_by_id,
            decisions_by_line,
        )
        original_lines = self._original_plate_resolve_lines(metadata, current_text)
        merged_lines = self._rewrite_plate_resolve_lines(
            spec,
            original_lines,
            items,
            resolved_by_line,
            resolved_decisions,
            plate_order_ctx=plate_order_ctx,
            for_batches=False,
        )
        if not merged_lines:
            raise ValueError(spec.empty_after_error)

        next_text = "\n".join(merged_lines)
        updated_batches = self._apply_plate_resolve_decisions_to_batches(
            spec,
            list(metadata.get("plate_batches") or []),
            items,
            resolved_by_line,
            resolved_decisions,
            plate_order_ctx=plate_order_ctx,
        )
        return self._persist_resolved_plate_preview(
            draft_id,
            metadata=metadata,
            next_text=next_text,
            updated_batches=updated_batches,
            resolved_decisions=resolved_decisions,
            plate_order_ctx=plate_order_ctx,
            spec=spec,
        )

    def _normalize_wide_plate_lines(self, items: Iterable[Any]) -> list[dict[str, Any]]:
        return self._wf.draft_service.serialize_wide_plate_lines(items)

    def _normalize_replacement_lines(
        self,
        replacement_text: str,
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> list[str]:
        if not replacement_text.strip():
            return []
        preview = self._wf.commercial_service.generate_preview(
            text=replacement_text,
            plate_order_ctx=plate_order_ctx,
        )
        return list(preview.parse_result.normalized_lines)

    def _load_plate_resolve_items(
        self,
        spec: PlateResolveSpec,
        metadata: dict[str, Any],
    ) -> list[dict[str, Any]]:
        if spec.kind == "wide":
            return self._normalize_wide_plate_lines(metadata.get("wide_plate_lines", []))
        if spec.kind == "invalid_width":
            return self._wf.draft_service.serialize_invalid_width_lines(
                metadata.get("invalid_width_lines", [])
            )
        return self._wf.draft_service.serialize_unpriced_plate_lines(
            metadata.get("unpriced_plate_lines", [])
        )

    @staticmethod
    def _index_plate_resolve_decisions(
        decisions: Iterable[dict[str, Any]],
        *,
        coerce_blank_source_line: bool,
    ) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
        decisions_by_id: dict[str, dict[str, Any]] = {}
        decisions_by_line: dict[str, dict[str, Any]] = {}
        for item in decisions:
            line_id = str(item.get("line_id", "") or "").strip()
            raw_source = item.get("source_line", "")
            source_line = (
                str(raw_source or "").strip()
                if coerce_blank_source_line
                else str(raw_source).strip()
            )
            if line_id:
                decisions_by_id[line_id] = item
            if source_line:
                decisions_by_line[source_line] = item
        return decisions_by_id, decisions_by_line

    def _bind_plate_resolve_items(
        self,
        spec: PlateResolveSpec,
        items: list[dict[str, Any]],
        decisions_by_id: dict[str, dict[str, Any]],
        decisions_by_line: dict[str, dict[str, Any]],
    ) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
        resolved_by_line: dict[str, dict[str, Any]] = {}
        resolved_decisions: dict[str, dict[str, Any]] = {}
        unresolved: list[str] = []
        for item in items:
            line_id = str(item.get("id", "")).strip()
            line = str(item.get("line", "")).strip()
            decision = decisions_by_id.get(line_id) if line_id else None
            if decision is None:
                decision = decisions_by_line.get(line)
            if decision is None:
                unresolved.append(line or line_id)
                continue
            stored = self._bind_plate_resolve_decision(spec, item, decision)
            resolved_decisions[line_id or line] = stored
            if line:
                resolved_by_line[line] = stored
        if unresolved:
            raise ValueError(spec.unresolved_error)
        return resolved_by_line, resolved_decisions

    def _bind_plate_resolve_decision(
        self,
        spec: PlateResolveSpec,
        item: dict[str, Any],
        decision: dict[str, Any],
    ) -> dict[str, Any]:
        if spec.kind == "wide":
            return decision
        if spec.kind == "invalid_width":
            return self._bind_invalid_width_decision(item, decision)
        return self._bind_unpriced_plate_decision(item, decision)

    @staticmethod
    def _bind_invalid_width_decision(
        item: dict[str, Any],
        decision: dict[str, Any],
    ) -> dict[str, Any]:
        action = str(decision.get("action", "")).strip().lower()
        allowed_widths = {
            int(repl["width_mm"])
            for repl in (item.get("replacements") or [])
            if isinstance(repl, dict) and repl.get("width_mm") is not None
        }
        if action == "replace_width":
            if not allowed_widths:
                raise ValueError(
                    "Для позиции без заводских замен доступно только исключение."
                )
            raw_width = decision.get("width_mm")
            if raw_width is None:
                raise ValueError("Для замены ширины нужно указать width_mm.")
            try:
                chosen_width = int(raw_width)
            except (TypeError, ValueError) as exc:
                raise ValueError("Некорректный width_mm для замены ширины.") from exc
            if chosen_width not in allowed_widths:
                raise ValueError(
                    f"width_mm={chosen_width} не входит в предложенные замены."
                )
        elif action == "exclude":
            pass
        else:
            raise ValueError("Некорректное действие для позиции с нестандартной шириной.")

        line_id = str(item.get("id", "")).strip()
        line = str(item.get("line", "")).strip()
        return {
            "line_id": line_id or None,
            "source_line": line or None,
            "action": action,
            "width_mm": decision.get("width_mm"),
        }

    @staticmethod
    def _bind_unpriced_plate_decision(
        item: dict[str, Any],
        decision: dict[str, Any],
    ) -> dict[str, Any]:
        action = str(decision.get("action", "")).strip().lower()
        allowed_load_codes = {
            int(repl["load_code"])
            for repl in (item.get("replacements") or [])
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

        line_id = str(item.get("id", "")).strip()
        line = str(item.get("line", "")).strip()
        return {
            "line_id": line_id or None,
            "source_line": line or None,
            "action": action,
            "load_code": decision.get("load_code"),
        }

    @staticmethod
    def _original_plate_resolve_lines(
        metadata: dict[str, Any],
        current_text: str,
    ) -> list[str]:
        original_lines = [
            line.strip() for line in list(metadata.get("normalized_lines") or []) if line.strip()
        ]
        if not original_lines:
            original_lines = [
                line.strip() for line in re.split(r"[\n;]+", current_text) if line.strip()
            ]
        return original_lines

    def _rewrite_plate_resolve_lines(
        self,
        spec: PlateResolveSpec,
        original_lines: list[str],
        items: list[dict[str, Any]],
        resolved_by_line: dict[str, dict[str, Any]],
        resolved_decisions: dict[str, dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
        for_batches: bool,
    ) -> list[str]:
        wide_lines: set[str] | None = None
        if spec.kind == "wide":
            wide_lines = {
                str(item.get("line", "")).strip()
                for item in items
                if str(item.get("line", "")).strip()
            }
        merged_lines: list[str] = []
        for line in original_lines:
            item, decision = self._lookup_plate_resolve_decision(
                spec,
                line,
                items,
                resolved_by_line,
                resolved_decisions,
                wide_lines=wide_lines,
                for_batches=for_batches,
            )
            if decision is None:
                merged_lines.append(line)
                continue
            merged_lines.extend(
                self._apply_plate_resolve_action(
                    spec,
                    line,
                    item,
                    decision,
                    plate_order_ctx=plate_order_ctx,
                    for_batches=for_batches,
                )
            )
        return merged_lines

    @staticmethod
    def _lookup_plate_resolve_decision(
        spec: PlateResolveSpec,
        line: str,
        items: list[dict[str, Any]],
        resolved_by_line: dict[str, dict[str, Any]],
        resolved_decisions: dict[str, dict[str, Any]],
        *,
        wide_lines: set[str] | None,
        for_batches: bool,
    ) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
        if spec.kind == "wide":
            if wide_lines is None or line not in wide_lines:
                return None, None
            if for_batches:
                return None, resolved_by_line.get(line)
            return None, resolved_by_line[line]

        matched_item = next(
            (
                item
                for item in items
                if str(item.get("line", "")).strip() == line
                or (
                    str(item.get("name", "")).strip()
                    and str(item.get("name", "")).strip() in line
                )
            ),
            None,
        )
        if matched_item is None:
            return None, None
        item_line = str(matched_item.get("line", "")).strip()
        item_id = str(matched_item.get("id", "")).strip()
        decision = (
            resolved_by_line.get(item_line)
            or resolved_decisions.get(item_id)
            or resolved_decisions.get(item_line)
        )
        return matched_item, decision

    def _apply_plate_resolve_action(
        self,
        spec: PlateResolveSpec,
        line: str,
        item: dict[str, Any] | None,
        decision: dict[str, Any],
        *,
        plate_order_ctx: PlateOrderContext,
        for_batches: bool,
    ) -> list[str]:
        action = str(decision.get("action", "")).strip().lower()
        if spec.kind == "wide":
            if action == "confirm":
                return [line]
            if action == "exclude":
                return []
            if action == "replace":
                replacement_text = str(decision.get("replacement_text", "") or "").strip()
                replacement_lines = self._normalize_replacement_lines(
                    replacement_text,
                    plate_order_ctx=plate_order_ctx,
                )
                if not replacement_lines and not for_batches:
                    raise ValueError(
                        "Для замены широкой плиты нужно указать корректный список замен."
                    )
                return replacement_lines
            if for_batches:
                return [line]
            raise ValueError(spec.invalid_action_error)

        if action == "exclude":
            return []
        if spec.kind == "invalid_width" and action == "replace_width":
            new_width = int(decision["width_mm"])
            rewritten = self._rewrite_invalid_width_line(
                line,
                item or {},
                new_width,
                restore_qty=not for_batches,
            )
            return [rewritten]
        if action == "replace_load":
            new_load = int(decision["load_code"])
            rewritten = self._rewrite_unpriced_load_line(
                line,
                item or {},
                new_load,
                restore_qty=not for_batches,
            )
            return [rewritten]
        if for_batches:
            return [line]
        raise ValueError(spec.invalid_action_error)

    @staticmethod
    def _rewrite_invalid_width_line(
        line: str,
        item: dict[str, Any],
        new_width_mm: int,
        *,
        restore_qty: bool,
    ) -> str:
        from core.factory_width import rewrite_plate_line_width

        try:
            return rewrite_plate_line_width(line, new_width_mm)
        except ValueError:
            fallback = str(item.get("name") or line)
            rewritten = rewrite_plate_line_width(fallback, new_width_mm)
            if restore_qty:
                qty_match = re.search(r"(\d+)\s*$", line.strip())
                if qty_match and not re.search(r"\d+\s*$", rewritten.strip()):
                    rewritten = f"{rewritten} {qty_match.group(1)}"
            return rewritten

    @staticmethod
    def _rewrite_unpriced_load_line(
        line: str,
        item: dict[str, Any],
        new_load: int,
        *,
        restore_qty: bool,
    ) -> str:
        from core.unpriced_plate_replacements import rewrite_plate_line_load

        try:
            return rewrite_plate_line_load(line, new_load)
        except ValueError:
            fallback = str(item.get("name") or line)
            rewritten = rewrite_plate_line_load(fallback, new_load)
            if restore_qty:
                qty_match = re.search(r"(\d+)\s*$", line.strip())
                if qty_match and not re.search(r"\d+\s*$", rewritten.strip()):
                    rewritten = f"{rewritten} {qty_match.group(1)}"
            return rewritten

    def _apply_plate_resolve_decisions_to_batches(
        self,
        spec: PlateResolveSpec,
        plate_batches: list[dict[str, Any]],
        items: list[dict[str, Any]],
        resolved_by_line: dict[str, dict[str, Any]],
        resolved_decisions: dict[str, dict[str, Any]],
        *,
        plate_order_ctx: PlateOrderContext,
    ) -> list[dict[str, Any]]:
        if not plate_batches:
            return plate_batches
        updated: list[dict[str, Any]] = []
        for batch in plate_batches:
            batch_text = str(batch.get("normalized_text", "") or "")
            batch_lines = [line.strip() for line in batch_text.split("\n") if line.strip()]
            next_lines = self._rewrite_plate_resolve_lines(
                spec,
                batch_lines,
                items,
                resolved_by_line,
                resolved_decisions,
                plate_order_ctx=plate_order_ctx,
                for_batches=True,
            )
            next_batch = dict(batch)
            next_batch["normalized_text"] = "\n".join(next_lines)
            updated.append(next_batch)
        return updated

    def _persist_resolved_plate_preview(
        self,
        draft_id: str,
        *,
        metadata: dict[str, Any],
        next_text: str,
        updated_batches: list[dict[str, Any]],
        resolved_decisions: dict[str, dict[str, Any]],
        plate_order_ctx: PlateOrderContext,
        spec: PlateResolveSpec,
    ) -> dict[str, Any]:
        preview = self._wf.commercial_service.generate_preview(
            text=next_text,
            plate_order_ctx=plate_order_ctx,
        )
        wide_resolved = (
            True
            if spec.force_wide_resolved
            else bool(metadata.get("wide_plates_resolved", True))
        )
        next_metadata = self._wf.draft_service.build_preview_metadata(
            preview=preview,
            base_metadata=metadata,
            source_type=str(metadata.get("source_type") or "text"),
            original_text=str(metadata.get("original_text", "") or ""),
            ocr_text=str(metadata.get("ocr_text", "") or ""),
            input_text=next_text,
            last_source_filename=str(metadata.get("last_source_filename", "") or ""),
            plate_batches=updated_batches,
            wide_plates_resolved=wide_resolved,
            source_metadata={},
        )
        if spec.extra_resolved_flag:
            next_metadata[spec.extra_resolved_flag] = True
        next_metadata[spec.decisions_meta_key] = list(resolved_decisions.values())
        order_data = self._wf._stamp_order_data(preview.order_data, product_type="plates")
        self._wf.draft_store.replace_preview(
            draft_id,
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=order_data,
            metadata=next_metadata,
        )
        self._wf._persist_wizard_step(draft_id, WizardStepId.plates)
        return self._wf.get_draft_details(draft_id)
