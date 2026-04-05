from __future__ import annotations

import json
import uuid
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any

from app.core.settings import get_settings
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder


class DraftStore:
    def __init__(self) -> None:
        self.settings = get_settings()
        self.base_dir = Path(self.settings.drafts_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def save_preview(
        self,
        *,
        order: PlateOrder,
        optimization_context: OptimizationContext,
        order_data: list[dict[str, Any]],
        metadata: dict[str, Any] | None = None,
    ) -> str:
        draft_id = uuid.uuid4().hex
        payload = {
            "order": order.to_dict(),
            "optimization": {
                "optimization_result": self._to_jsonable(optimization_context.optimization_result),
                "plan_by_load": self._to_jsonable(optimization_context.plan_by_load),
                "load_to_reinforcement_map": self._to_jsonable(optimization_context.load_to_reinforcement_map),
            },
            "order_data": self._to_jsonable(order_data),
            "metadata": self._to_jsonable(metadata or {}),
        }
        with open(self.base_dir / f"{draft_id}.json", "w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)
        return draft_id

    def load_preview(self, draft_id: str) -> dict[str, Any] | None:
        path = self.base_dir / f"{draft_id}.json"
        if not path.exists():
            return None
        with open(path, "r", encoding="utf-8") as file:
            payload = json.load(file)
        payload["order"] = PlateOrder.from_dict(payload["order"])
        optimization_payload = payload.get("optimization", {})
        payload["optimization_context"] = OptimizationContext(
            order=payload["order"],
            optimization_result=optimization_payload.get("optimization_result", {}),
            plan_by_load=optimization_payload.get("plan_by_load", {}),
            load_to_reinforcement_map=optimization_payload.get("load_to_reinforcement_map", {}),
        )
        return payload

    def _to_jsonable(self, value: Any) -> Any:
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        if is_dataclass(value):
            return self._to_jsonable(asdict(value))
        if isinstance(value, dict):
            normalized: dict[str, Any] = {}
            for key, item in value.items():
                normalized[str(key)] = self._to_jsonable(item)
            return normalized
        if isinstance(value, (list, tuple, set)):
            return [self._to_jsonable(item) for item in value]
        if hasattr(value, "to_dict") and callable(value.to_dict):
            return self._to_jsonable(value.to_dict())
        if hasattr(value, "model_dump") and callable(value.model_dump):
            return self._to_jsonable(value.model_dump())
        if hasattr(value, "__dict__"):
            return {
                "__class__": value.__class__.__name__,
                "state": self._to_jsonable(vars(value)),
            }
        return str(value)

