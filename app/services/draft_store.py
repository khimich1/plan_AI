from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Any

from app.core.settings import get_settings
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder


def _json_safe_optimization_result(result: dict[str, Any]) -> dict[str, Any]:
    """Копия optimization_result без объектов, которые json.dump не сериализует (например PlateAudit)."""
    safe = dict(result)
    safe.pop("_plate_audit", None)
    return safe


def _json_safe_plan_by_load(plan_by_load: dict[str, Any]) -> dict[str, Any]:
    """То же для plan_by_load: значения — те же dict'ы, что и optimization_result (в т.ч. с _plate_audit)."""
    out: dict[str, Any] = {}
    for key, value in plan_by_load.items():
        if isinstance(value, dict):
            out[key] = _json_safe_optimization_result(value)
        else:
            out[key] = value
    return out


class DraftStore:
    def __init__(self) -> None:
        self.settings = get_settings()
        self.base_dir = Path(self.settings.drafts_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def _get_path(self, draft_id: str) -> Path:
        return self.base_dir / f"{draft_id}.json"

    def _load_raw(self, draft_id: str) -> dict[str, Any] | None:
        path = self._get_path(draft_id)
        if not path.exists():
            return None
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)

    def save_preview(
        self,
        *,
        order: PlateOrder,
        optimization_context: OptimizationContext,
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any] | None = None,
    ) -> str:
        draft_id = uuid.uuid4().hex
        self.replace_preview(
            draft_id,
            order=order,
            optimization_context=optimization_context,
            order_data=order_data,
            metadata=metadata,
        )
        return draft_id

    def replace_preview(
        self,
        draft_id: str,
        *,
        order: PlateOrder,
        optimization_context: OptimizationContext,
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any] | None = None,
    ) -> str:
        payload = {
            "order": order.to_dict(),
            "optimization": {
                "optimization_result": _json_safe_optimization_result(
                    optimization_context.optimization_result
                ),
                "plan_by_load": _json_safe_plan_by_load(optimization_context.plan_by_load),
                "load_to_reinforcement_map": optimization_context.load_to_reinforcement_map,
            },
            "order_data": order_data,
            "metadata": metadata or {},
        }
        with open(self._get_path(draft_id), "w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)
        return draft_id

    def load_preview(self, draft_id: str) -> dict[str, Any] | None:
        payload = self._load_raw(draft_id)
        if payload is None:
            return None
        payload["order"] = PlateOrder.from_dict(payload["order"])
        optimization_payload = payload.get("optimization", {})
        payload["optimization_context"] = OptimizationContext(
            order=payload["order"],
            optimization_result=optimization_payload.get("optimization_result", {}),
            plan_by_load=optimization_payload.get("plan_by_load", {}),
            load_to_reinforcement_map=optimization_payload.get("load_to_reinforcement_map", {}),
        )
        return payload

    def update_metadata(self, draft_id: str, **updates: Any) -> dict[str, Any] | None:
        payload = self._load_raw(draft_id)
        if payload is None:
            return None
        metadata = dict(payload.get("metadata", {}))
        metadata.update(updates)
        payload["metadata"] = metadata
        with open(self._get_path(draft_id), "w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)
        return payload

    def delete(self, draft_id: str) -> None:
        self._get_path(draft_id).unlink(missing_ok=True)

