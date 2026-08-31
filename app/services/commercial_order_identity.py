"""Stamp / partition / compose helpers for commercial order lines.

No workflow import: this module must stay free of CommercialWorkflowService.
"""

from __future__ import annotations

from typing import Any
from uuid import uuid4

from app.services.commercial_draft_service import CommercialDraftService
from app.services.product_draft_config import SPECS

APPEND_PRODUCT_TYPES = frozenset(SPECS)

PRODUCT_KIND_TO_TYPE: dict[str, str] = {
    "march": "marches",
    "step": "steps",
    "pile": "piles",
    "bridge_pile": "bridge_piles",
    "fbs": "fbs",
    "plate": "plates",
}


class CommercialOrderIdentity:
    def __init__(self, draft_service: CommercialDraftService) -> None:
        self._draft_service = draft_service

    def stamp_order_data(
        self,
        order_data: list[Any] | None,
        *,
        product_type: str,
        previous_order_data: list[Any] | None = None,
    ) -> list[dict[str, Any]]:
        return self._draft_service.stamp_order_line_identity(
            list(order_data or []),
            product_type=product_type,
            previous_order_data=list(previous_order_data or []) if previous_order_data else None,
        )

    @staticmethod
    def line_product_type(line: dict[str, Any] | None) -> str:
        if not isinstance(line, dict):
            return ""
        explicit = str(line.get("product_type") or "").strip().lower()
        if explicit:
            return explicit
        kind = str(line.get("product_kind") or "").strip().lower()
        return PRODUCT_KIND_TO_TYPE.get(kind, "")

    @staticmethod
    def line_is_sealed(line: dict[str, Any] | None) -> bool:
        """Sealed lines carry append_batch_id (assigned by ``seal_unbatched_lines``)."""
        if not isinstance(line, dict):
            return False
        return bool(str(line.get("append_batch_id") or "").strip())

    def partition_order_by_product_type(
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
            (t := self.line_product_type(line)) and t != normalized for line in rows
        )
        others: list[dict[str, Any]] = []
        same: list[dict[str, Any]] = []
        for line in rows:
            line_type = self.line_product_type(line)
            if line_type == normalized or (not line_type and not typed_conflict):
                same.append(line)
            else:
                others.append(line)
        return others, same

    def stamp_previous_for_product_update(
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
            return [ln for ln in same_previous if not self.line_is_sealed(ln)]
        if normalized_mode == "replace":
            return list(same_previous)
        return []

    def compose_order_data_for_product_update(
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
                    self.line_product_type(line) == normalized_type
                    and not self.line_is_sealed(line)
                )
                if not unsealed_same_type:
                    kept.append(line)
            return kept + list(new_type_lines)
        others, _same_previous = self.partition_order_by_product_type(
            previous_order_data,
            product_type=product_type,
        )
        return others + list(new_type_lines)

    def current_cycle_lines(
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
            if self.line_is_sealed(raw):
                continue
            line_type = self.line_product_type(raw)
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
            if isinstance(raw, dict) and not self.line_is_sealed(raw)
        ]
        if unsealed and all(not self.line_product_type(line) for line in unsealed):
            return unsealed
        return []

    def seal_unbatched_lines(
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
        if product_type not in APPEND_PRODUCT_TYPES:
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
