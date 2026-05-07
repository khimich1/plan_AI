from __future__ import annotations

import json
import logging
import os
import re
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from filelock import FileLock, Timeout

from app.core.settings import get_settings
from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.plate_order import PlateOrder
from core.serialization import (
    strip_plate_audit,
    strip_plate_audit_from_plan_by_load,
)

logger = logging.getLogger(__name__)

_DRAFT_ID_SAFE = re.compile(r"^[A-Za-z0-9._-]{1,128}$")


class UnsafeDraftIdError(ValueError):
    """Draft id is empty, malformed, or attempts to escape the drafts directory."""


class DraftStoreLockTimeout(TimeoutError):
    """Another process holds the draft lock longer than the configured timeout."""


class DraftStore:
    """Stores commercial offer drafts as JSON files under ``settings.drafts_dir``.

    For several API replicas, point ``DRAFTS_DIR`` (and ``OUTPUTS_DIR`` for generated
    files) at a single shared filesystem and set ``APP_STORAGE_LAYOUT=shared_volume``.
    Per-draft file locks coordinate concurrent updates on that shared volume.
    """

    def __init__(self) -> None:
        self.settings = get_settings()
        self.base_dir = Path(self.settings.drafts_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self._base_resolved = self.base_dir.resolve()
        self._lock_timeout = float(self.settings.draft_store_lock_timeout_seconds)

    @staticmethod
    def validate_draft_id(draft_id: str) -> None:
        if not draft_id or not isinstance(draft_id, str):
            raise UnsafeDraftIdError("Invalid draft id.")
        stripped = draft_id.strip()
        if stripped != draft_id:
            raise UnsafeDraftIdError("Invalid draft id.")
        if not _DRAFT_ID_SAFE.match(draft_id):
            raise UnsafeDraftIdError("Invalid draft id.")
        if ".." in draft_id:
            raise UnsafeDraftIdError("Invalid draft id.")
        if Path(draft_id).is_absolute():
            raise UnsafeDraftIdError("Invalid draft id.")

    def _get_path(self, draft_id: str) -> Path:
        self.validate_draft_id(draft_id)
        candidate = (self.base_dir / f"{draft_id}.json").resolve()
        try:
            candidate.relative_to(self._base_resolved)
        except ValueError:
            logger.warning(
                "Rejected draft path outside drafts_dir (draft_id prefix=%r)",
                draft_id[:16],
            )
            raise UnsafeDraftIdError("Invalid draft id.") from None
        return candidate

    def _lock_path(self, draft_id: str) -> Path:
        self.validate_draft_id(draft_id)
        candidate = (self.base_dir / f"{draft_id}.lock").resolve()
        try:
            candidate.relative_to(self._base_resolved)
        except ValueError:
            raise UnsafeDraftIdError("Invalid draft id.") from None
        return candidate

    @contextmanager
    def _draft_lock(self, draft_id: str) -> Iterator[None]:
        lock = FileLock(str(self._lock_path(draft_id)), timeout=self._lock_timeout)
        try:
            with lock:
                yield
        except Timeout as exc:
            logger.warning(
                "Draft lock timeout (prefix=%r, timeout=%s)",
                draft_id[:16],
                self._lock_timeout,
            )
            raise DraftStoreLockTimeout(
                f"Could not acquire draft lock within {self._lock_timeout} seconds."
            ) from exc

    def _atomic_write_json(self, path: Path, data: dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = path.with_suffix(path.suffix + ".tmp")
        try:
            with open(tmp_path, "w", encoding="utf-8") as file:
                json.dump(data, file, ensure_ascii=False, indent=2)
                file.flush()
                os.fsync(file.fileno())
            os.replace(tmp_path, path)
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass

    def _load_raw_nolock(self, draft_id: str) -> dict[str, Any] | None:
        path = self._get_path(draft_id)
        if not path.exists():
            return None
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)

    def load_raw_json(self, draft_id: str) -> dict[str, Any] | None:
        """Load draft JSON without deserializing domain models (ownership checks, guards)."""
        with self._draft_lock(draft_id):
            return self._load_raw_nolock(draft_id)

    def generated_files_filenames(self, draft_id: str) -> set[str]:
        """Basenames of files recorded for this draft (for download authorization)."""
        with self._draft_lock(draft_id):
            raw = self._load_raw_nolock(draft_id)
        if raw is None:
            return set()
        meta = raw.get("metadata") or {}
        names: set[str] = set()
        for item in meta.get("generated_files") or []:
            name = Path(str(item.get("filename", "")).strip()).name
            if name:
                names.add(name)
        return names

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
                "optimization_result": strip_plate_audit(
                    optimization_context.optimization_result
                ),
                "plan_by_load": strip_plate_audit_from_plan_by_load(
                    optimization_context.plan_by_load
                ),
                "load_to_reinforcement_map": optimization_context.load_to_reinforcement_map,
            },
            "order_data": order_data,
            "metadata": metadata or {},
        }
        path = self._get_path(draft_id)
        with self._draft_lock(draft_id):
            self._atomic_write_json(path, payload)
        return draft_id

    def load_preview(self, draft_id: str) -> dict[str, Any] | None:
        with self._draft_lock(draft_id):
            payload = self._load_raw_nolock(draft_id)
        if payload is None:
            return None
        payload["order"] = PlateOrder.from_dict(payload["order"])
        optimization_payload = payload.get("optimization", {})
        payload["optimization_context"] = OptimizationContext(
            order=payload["order"],
            optimization_result=optimization_payload.get("optimization_result", {}),
            plan_by_load=optimization_payload.get("plan_by_load", {}),
            load_to_reinforcement_map=optimization_payload.get(
                "load_to_reinforcement_map", {}
            ),
        )
        return payload

    def update_metadata(self, draft_id: str, **updates: Any) -> dict[str, Any] | None:
        with self._draft_lock(draft_id):
            payload = self._load_raw_nolock(draft_id)
            if payload is None:
                return None
            metadata = dict(payload.get("metadata", {}))
            metadata.update(updates)
            payload["metadata"] = metadata
            self._atomic_write_json(self._get_path(draft_id), payload)
        return payload

    def delete(self, draft_id: str) -> None:
        path = self._get_path(draft_id)
        with self._draft_lock(draft_id):
            path.unlink(missing_ok=True)
