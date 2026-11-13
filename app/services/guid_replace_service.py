"""Замена марки экономистом во всех КП открытой строки очереди.

Сначала все подходящие КП одним соединением и одним commit, затем выбор
дубля (если карточек две), затем сверка очереди. Сверка наружу не бросает.
"""

from __future__ import annotations

import logging
import sqlite3
from dataclasses import dataclass

from core.domain.enums import PlateStatus
from core.guid_demand import list_open
from core.guid_gate import (
    REASON_AMBIGUOUS,
    REASON_DUP,
    OrderLine,
    check_invoice_guids,
)
from core.guid_replace import (
    GuidReplaceError,
    ReplacementPreview,
    ReplaceTarget,
    _line_mark,
    build_replacement,
    preview_replacement,
)
from core.kp.offers_read import get_kp_by_id
from core.kp.offers_write import update_kp_from_order_data
from core.kp_db_common import _connect
from core.kp_order_data import order_data_from_kp_info

logger = logging.getLogger(__name__)

QUEUE_NOTICE = "Очередь обновится при открытии страницы."
_ARCHIVED = "в архиве"
_PROTECTED_PLATE_STATUSES = frozenset(
    {PlateStatus.IN_PLAN.value, PlateStatus.IN_PRODUCTION.value}
)
_PLATE_LOCK_TEXT = "Плита уже в плане или в производстве, замените вручную."


@dataclass(frozen=True)
class SkippedKp:
    kp_id: int
    reason: str


@dataclass(frozen=True)
class WrittenKp:
    kp_id: int
    total: float


@dataclass(frozen=True)
class ReplaceQueueResult:
    written: tuple[WrittenKp, ...]
    skipped: tuple[SkippedKp, ...]
    queue_notice: str | None


@dataclass(frozen=True)
class QueuePreviewItem:
    kp_id: int
    preview: ReplacementPreview


@dataclass(frozen=True)
class PreviewQueueResult:
    items: tuple[QueuePreviewItem, ...]
    skipped: tuple[SkippedKp, ...]


@dataclass
class _ReadyKp:
    kp_id: int
    order: list[dict]
    discount: float
    logistics: float


class GuidReplaceService:
    def __init__(self, plita_db_path: str, pb_db_path: str) -> None:
        self.plita_db_path = str(plita_db_path)
        self.pb_db_path = str(pb_db_path)

    def preview_in_queue(
        self,
        *,
        product_kind: str,
        mark: str,
        target_mark: str,
        nomenclature_id: str | None = None,
        replace_oneoff: bool = False,
    ) -> PreviewQueueResult:
        ready, skipped, _chosen_guid = self._prepare(
            product_kind=product_kind,
            mark=mark,
            target_mark=target_mark,
            nomenclature_id=nomenclature_id,
            replace_oneoff=replace_oneoff,
        )
        target = ReplaceTarget(target_mark, _chosen_guid)
        items = tuple(
            QueuePreviewItem(
                kp_id=item.kp_id,
                preview=preview_replacement(
                    item.order,
                    source_mark=mark,
                    target=target,
                    discount_percent=item.discount,
                    kind=product_kind,
                    actor="economist",
                    pb_db_path=self.pb_db_path,
                    totals_db_path=self.plita_db_path,
                    kp_id=item.kp_id,
                    logistics_cost=item.logistics,
                ),
            )
            for item in ready
        )
        # preview_replacement строит новую строку внутри себя; для окна достаточно
        # уже посчитанных денег. Готовые order здесь — исходные, запись не идёт.
        return PreviewQueueResult(items=items, skipped=tuple(skipped))

    def replace_in_queue(
        self,
        *,
        product_kind: str,
        mark: str,
        target_mark: str,
        nomenclature_id: str | None = None,
        changed_by: str | None = None,
        replace_oneoff: bool = False,
    ) -> ReplaceQueueResult:
        ready, skipped, chosen_guid = self._prepare(
            product_kind=product_kind,
            mark=mark,
            target_mark=target_mark,
            nomenclature_id=nomenclature_id,
            replace_oneoff=replace_oneoff,
        )
        if not ready:
            return ReplaceQueueResult((), tuple(skipped), None)
        new_orders = self._replaced_orders(
            ready,
            product_kind=product_kind,
            mark=mark,
            target_mark=target_mark,
            nomenclature_id=chosen_guid,
            replace_oneoff=replace_oneoff,
        )
        self._write_all(new_orders, changed_by=changed_by)
        notice = self._finish_queue(
            product_kind=product_kind,
            target_mark=target_mark,
            nomenclature_id=nomenclature_id,
            changed_by=changed_by,
        )
        written = tuple(
            WrittenKp(kp_id=item.kp_id, total=self._total(item.kp_id)) for item in ready
        )
        return ReplaceQueueResult(written, tuple(skipped), notice)

    def _prepare(
        self,
        *,
        product_kind: str,
        mark: str,
        target_mark: str,
        nomenclature_id: str | None,
        replace_oneoff: bool,
    ) -> tuple[list[_ReadyKp], list[SkippedKp], str | None]:
        kind = str(product_kind or "").strip()
        source = str(mark or "").strip()
        chosen = str(target_mark or "").strip()
        if not kind or not source or not chosen:
            raise GuidReplaceError("Нужны вид, марка очереди и выбранная марка.")
        kp_ids = self._open_kp_ids(kind, source)
        if not kp_ids:
            raise GuidReplaceError(f"В очереди нет открытой строки «{source}».")
        self._require_choice_when_two(kind, chosen, nomenclature_id)
        chosen_guid = self._fill_single_guid(kind, chosen, nomenclature_id)
        ready: list[_ReadyKp] = []
        skipped: list[SkippedKp] = []
        target = ReplaceTarget(chosen, chosen_guid)
        for kp_id in kp_ids:
            decision = self._one_kp(
                kp_id,
                kind=kind,
                source=source,
                target=target,
                replace_oneoff=replace_oneoff,
            )
            if isinstance(decision, SkippedKp):
                skipped.append(decision)
            else:
                ready.append(decision)
        return ready, skipped, chosen_guid

    def _one_kp(
        self,
        kp_id: int,
        *,
        kind: str,
        source: str,
        target: ReplaceTarget,
        replace_oneoff: bool,
    ) -> _ReadyKp | SkippedKp:
        info = get_kp_by_id(kp_id, self.plita_db_path)
        if info is None:
            return SkippedKp(kp_id, f"КП №{kp_id} не найдено.")
        status = str(info.get("status") or "").strip()
        if status != _ARCHIVED:
            shown = status or "без статуса"
            return SkippedKp(
                kp_id,
                f"КП №{kp_id} в статусе «{shown}»: замена пишет только КП «в архиве».",
            )
        order = order_data_from_kp_info(info)
        self._stamp_plate_grades(info, order)
        matched = [line for line in order if _line_mark(line, kind) == source]
        if not matched:
            return SkippedKp(kp_id, f"В КП №{kp_id} нет марки «{source}».")
        if any(not str(line.get("line_id") or "").strip() for line in matched):
            return SkippedKp(kp_id, f"В КП №{kp_id} у строки нет line_id, замена невозможна.")
        if kind == "plate" and self._plate_is_locked(info, matched):
            return SkippedKp(kp_id, f"КП №{kp_id}: {_PLATE_LOCK_TEXT}")
        try:
            for line in matched:
                build_replacement(
                    line,
                    target,
                    float(info.get("discount_percent") or 0.0),
                    kind,
                    actor="economist",
                    pb_db_path=self.pb_db_path,
                    replace_oneoff=replace_oneoff,
                )
        except GuidReplaceError as exc:
            return SkippedKp(kp_id, f"КП №{kp_id}: {exc.detail}")
        return _ReadyKp(
            kp_id=kp_id,
            order=order,
            discount=float(info.get("discount_percent") or 0.0),
            logistics=float(info.get("logistics_cost") or 0.0),
        )

    def _replaced_orders(
        self,
        ready: list[_ReadyKp],
        *,
        product_kind: str,
        mark: str,
        target_mark: str,
        nomenclature_id: str | None,
        replace_oneoff: bool,
    ) -> list[tuple[_ReadyKp, list[dict]]]:
        target = ReplaceTarget(target_mark, nomenclature_id)
        built: list[tuple[_ReadyKp, list[dict]]] = []
        for item in ready:
            order = [dict(line) for line in item.order]
            for index, line in enumerate(order):
                if _line_mark(line, product_kind) != str(mark).strip():
                    continue
                order[index] = build_replacement(
                    line,
                    target,
                    item.discount,
                    product_kind,
                    actor="economist",
                    pb_db_path=self.pb_db_path,
                    replace_oneoff=replace_oneoff,
                )
            built.append((item, order))
        return built

    def _write_all(
        self,
        orders: list[tuple[_ReadyKp, list[dict]]],
        *,
        changed_by: str | None,
    ) -> None:
        conn = _connect(self.plita_db_path)
        try:
            for item, order in orders:
                update_kp_from_order_data(
                    item.kp_id,
                    order,
                    discount_percent=item.discount,
                    db_path=self.plita_db_path,
                    _external_conn=conn,
                    changed_by=changed_by,
                    history_source="replace_economist",
                )
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
        from app.services.commercial_draft_lifecycle import recalc_promise_after_edit

        for item, _order in orders:
            recalc_promise_after_edit(self.plita_db_path, item.kp_id)

    def _finish_queue(
        self,
        *,
        product_kind: str,
        target_mark: str,
        nomenclature_id: str | None,
        changed_by: str | None,
    ) -> str | None:
        failed = False
        if nomenclature_id and self._guid_count(product_kind, target_mark) >= 2:
            try:
                self._save_choice(
                    product_kind, target_mark, nomenclature_id, changed_by
                )
            except Exception:
                logger.exception("Не удалось записать выбор карточки после замены")
                failed = True
        try:
            from app.services.guid_queue_sync import sync_guid_queue_with_kp

            report = sync_guid_queue_with_kp(self.pb_db_path, self.plita_db_path)
        except Exception:
            logger.exception("Не удалось сверить очередь после замены марки")
            report = None
        if report is None:
            logger.error("Сверка очереди после замены не выполнена")
            failed = True
        return QUEUE_NOTICE if failed else None

    def _save_choice(
        self,
        product_kind: str,
        target_mark: str,
        nomenclature_id: str,
        changed_by: str | None,
    ) -> None:
        from core.guid_demand import sweep_guid_demand

        conn = sqlite3.connect(self.pb_db_path)
        try:
            if product_kind == "plate":
                from core.plate_guid_choice import set_choice

                set_choice(
                    conn,
                    target_mark,
                    chosen_guid=nomenclature_id,
                    chosen_name=target_mark,
                    decided_by=changed_by,
                    note="замена марки",
                )
            else:
                from core.nomenclature_guid import set_manual_choice

                set_manual_choice(
                    conn,
                    product_kind,
                    target_mark,
                    guid_1c=nomenclature_id,
                    decided_by=changed_by,
                    note="замена марки",
                )
            sweep_guid_demand(conn)
            conn.commit()
        finally:
            conn.close()

    def _open_kp_ids(self, kind: str, mark: str) -> list[int]:
        conn = sqlite3.connect(self.pb_db_path)
        try:
            ids: list[int] = []
            for row in list_open(conn):
                if row.product_kind == kind and row.mark == mark:
                    ids.extend(int(kp_id) for kp_id in row.kp_ids)
        finally:
            conn.close()
        return list(dict.fromkeys(ids))

    def _fill_single_guid(
        self, kind: str, mark: str, given: str | None
    ) -> str | None:
        text = str(given or "").strip()
        if text:
            return text
        if self._guid_count(kind, mark) != 1:
            return None
        conn = sqlite3.connect(self.pb_db_path)
        try:
            if kind == "plate":
                from core.kp_db_nomenclature import find_plate_card

                match = find_plate_card(conn.cursor(), mark)
                return match.guid if match.guid_count == 1 else None
            from core.nomenclature_guid import get_guid_for_invoice

            if kind == "pile" and mark.endswith("у"):
                return get_guid_for_invoice(
                    conn, kind, mark[:-1], reinforced=True
                ) or get_guid_for_invoice(conn, kind, mark)
            return get_guid_for_invoice(conn, kind, mark)
        finally:
            conn.close()

    def _require_choice_when_two(
        self, kind: str, target_mark: str, nomenclature_id: str | None
    ) -> None:
        if self._guid_count(kind, target_mark) >= 2 and not str(nomenclature_id or "").strip():
            raise GuidReplaceError("Выберите одну из двух карточек.")

    def _guid_count(self, kind: str, mark: str) -> int:
        conn = sqlite3.connect(self.pb_db_path)
        try:
            report = check_invoice_guids([OrderLine(kind, mark)], conn)
        finally:
            conn.close()
        if not report.missing:
            return 1
        if report.missing[0].reason in {REASON_DUP, REASON_AMBIGUOUS}:
            return 2
        return 0

    def _plate_is_locked(self, info: dict, matched: list[dict]) -> bool:
        wanted = {str(line.get("line_id") or "").strip() for line in matched}
        for plate in info.get("plates") or []:
            line_id = str(plate.get("line_id") or "").strip()
            if line_id not in wanted:
                continue
            status = str(plate.get("status") or "").strip()
            plan_id = str(plate.get("plan_id") or "").strip()
            if status in _PROTECTED_PLATE_STATUSES or plan_id:
                return True
        return False

    @staticmethod
    def _stamp_plate_grades(info: dict, order: list[dict]) -> None:
        by_line = {
            str(plate.get("line_id") or "").strip(): plate
            for plate in info.get("plates") or []
        }
        for line in order:
            if str(line.get("concrete_grade") or "").strip():
                continue
            src = by_line.get(str(line.get("line_id") or "").strip())
            if src is None:
                continue
            grade = str(src.get("concrete_grade") or "").strip()
            if grade:
                line["concrete_grade"] = grade

    def _total(self, kp_id: int) -> float:
        with sqlite3.connect(self.plita_db_path) as conn:
            row = conn.execute(
                "SELECT total_amount FROM KP_offers WHERE kp_id = ?",
                (kp_id,),
            ).fetchone()
        return float(row[0] or 0.0) if row else 0.0
