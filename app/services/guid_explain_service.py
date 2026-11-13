"""«Почему нет GUID»: чтение pb.db для менеджера (шаг 3) и экономиста (очередь)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

from app.core.settings import get_settings
from app.schemas.guid_explain import (
    GuidCandidateOut,
    GuidCardOut,
    GuidExplainResponse,
    GuidSearchResponse,
)
from core.guid_explain import (
    GuidCandidate,
    explain_missing_guid,
    search_guid_marks,
)
from core.guid_gate import OrderLine, _normalize_kind
from core.nomenclature_guid import PRODUCT_KINDS


class GuidExplainError(ValueError):
    """Неизвестный вид изделия или пустая марка."""


class GuidExplainService:
    def __init__(self, pb_db_path: Path | str | None = None) -> None:
        self._pb_db_path = Path(pb_db_path) if pb_db_path is not None else None

    def explain(self, product_kind: str, mark: str, *, role: str) -> GuidExplainResponse:
        kind = self._kind(product_kind)
        text = str(mark or "").strip()
        if not text:
            raise GuidExplainError("Укажите марку")
        conn = self._connect()
        try:
            result = explain_missing_guid(OrderLine(kind, text), conn)
        finally:
            conn.close()
        return GuidExplainResponse(
            product_kind=result.product_kind,
            mark=result.mark,
            reason=result.reason,
            hint=result.hint,
            same_cards=self._out(result.same_cards, role),
            slice=self._out(result.slice, role),
        )

    def search(self, product_kind: str, query: str, *, role: str) -> GuidSearchResponse:
        kind = self._kind(product_kind)
        conn = self._connect()
        try:
            found = search_guid_marks(conn, kind, query)
        finally:
            conn.close()
        return GuidSearchResponse(items=self._out(found, role))

    @staticmethod
    def _kind(product_kind: str) -> str:
        kind = _normalize_kind(product_kind)
        if kind != "plate" and kind not in PRODUCT_KINDS:
            raise GuidExplainError("Неизвестный вид изделия")
        return kind

    @staticmethod
    def _out(items: tuple[GuidCandidate, ...], role: str) -> list[GuidCandidateOut]:
        return [
            GuidCandidateOut(
                mark=item.mark,
                guid_count=item.guid_count,
                label=item.label,
                selectable=item.selectable_for(role),
                cards=[GuidCardOut(name=card.name, guid=card.guid) for card in item.cards],
            )
            for item in items
        ]

    def _connect(self) -> sqlite3.Connection:
        path = self._pb_db_path or get_settings().pb_db_path
        return sqlite3.connect(str(path))
