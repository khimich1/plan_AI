"""CRUD графика поставки + живой светофор (GET/PUT) + XLSX import/template + документ."""

from __future__ import annotations

import logging
import sqlite3
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Literal

from app.core.settings import get_settings
from app.domain.enums import KpStatus
from app.repositories.plan_repository import PlanRepository
from app.schemas.delivery_schedule import (
    BatchDraftItemOut,
    BatchDraftOut,
    BatchItemOut,
    BatchOut,
    DeliverySchedulePut,
    DeliveryScheduleView,
    ImportDraftResponse,
    UnmatchedRowOut,
)
from app.security.offer_access import assert_offer_read_access, assert_offer_write_access
from app.services.kp_readiness_service import KpReadinessService
from app.services.plan_distribution_service import PlanDistributionService
from core.delivery_schedule_check import BatchInput, BatchItemInput, check_batches
from core.delivery_schedule_pdf import build_document as build_pdf_document
from core.delivery_schedule_xlsx import (
    build_document as build_xlsx_document,
    build_template,
    parse_template,
)
from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema
from core.work_calendar import is_working_day, load_extra_workdays, load_holidays

logger = logging.getLogger(__name__)

ALLOWED_EDIT_KP_STATUSES: tuple[str, ...] = (
    KpStatus.IN_WORK.value,
    KpStatus.ON_SGP.value,
)

DocumentFmt = Literal["xlsx", "pdf"]

# Горизонт workdays для симуляции светофора (плюс запас за max produce_by).
_WORKDAYS_HORIZON_DAYS = 400
_WORKDAYS_AFTER_PRODUCE_BY = 60

# Ожидаемые сбои источников светофора (readiness / calendar / workdays / БД).
# Прочие исключения (TypeError и т.п.) не глотаем — пусть уходят в 500.
_TRAFFIC_LIGHT_SOURCE_ERRORS = (OSError, RuntimeError, sqlite3.Error, ConnectionError)


class DeliveryScheduleError(Exception):
    """Базовое исключение сервиса графика поставки."""


class DeliveryScheduleNotFoundError(DeliveryScheduleError):
    """График или КП не найдены."""


class DeliveryScheduleValidationError(DeliveryScheduleError):
    """Ошибка валидации бизнес-правил (маппится в 422)."""


class DeliveryScheduleService:
    """Полная замена партий (PUT) и чтение графика (GET) со светофором."""

    def __init__(
        self,
        *,
        db_path: str,
        outputs_dir: Path | None = None,
    ) -> None:
        settings = get_settings()
        self.db_path = db_path
        self.outputs_dir = Path(outputs_dir or settings.outputs_dir)
        self.outputs_dir.mkdir(parents=True, exist_ok=True)
        # Для детерминированных тестов: ISO YYYY-MM-DD.
        self._today_override: str | None = None

    def get(self, kp_id: int, *, user: dict, today: str | None = None) -> DeliveryScheduleView:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            offer = self._fetch_offer(cur, kp_id)
            if offer is None:
                raise DeliveryScheduleNotFoundError(f"КП №{kp_id} не найдено")
            assert_offer_read_access(user, offer)
            schedule = self._fetch_schedule(cur, kp_id)
            if schedule is None:
                raise DeliveryScheduleNotFoundError(
                    f"График поставки для КП №{kp_id} не найден"
                )
            view = self._build_view(cur, schedule)
        finally:
            conn.close()

        return self._enrich_with_traffic_light(view, kp_id, today=today)

    def replace(
        self,
        kp_id: int,
        payload: DeliverySchedulePut,
        user: dict,
        *,
        today: str | None = None,
    ) -> DeliveryScheduleView:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()

            offer = self._fetch_offer(cur, kp_id)
            if offer is None:
                raise DeliveryScheduleNotFoundError(f"КП №{kp_id} не найдено")
            assert_offer_write_access(user, offer)

            status = (offer.get("status") or "").strip()
            if status not in ALLOWED_EDIT_KP_STATUSES:
                raise DeliveryScheduleValidationError(
                    "График поставки можно редактировать только у КП "
                    "со статусом «в работе» или «На СГП»"
                )

            plate_qty = self._load_plate_qty_map(cur, kp_id)
            self._validate_batches_against_plates(payload, plate_qty)

            now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
            schedule_id = self._upsert_schedule_header(cur, kp_id, payload, now)
            self._replace_batches(cur, schedule_id, payload)
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

        return self.get(kp_id, user=user, today=today)

    def build_template_bytes(self, kp_id: int, *, user: dict) -> bytes:
        """XLSX-шаблон: сохранённые партии сверху, остаток КП снизу."""
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            offer = self._fetch_offer(cur, kp_id)
            if offer is None:
                raise DeliveryScheduleNotFoundError(f"КП №{kp_id} не найдено")
            assert_offer_read_access(user, offer)
            kp_plates = self._load_kp_plates_for_import(cur, kp_id)
            schedule = self._fetch_schedule(cur, kp_id)
            batches: list[dict] = []
            if schedule is not None:
                view = self._build_view(cur, schedule)
                batches = [
                    {
                        "name": batch.name,
                        "deliver_from": batch.deliver_from,
                        "deliver_to": batch.deliver_to,
                        "produce_by": batch.produce_by,
                        "items": [
                            {
                                "plate_id": item.plate_id,
                                "plate_name": item.plate_name,
                                "qty": item.qty,
                            }
                            for item in batch.items
                        ],
                    }
                    for batch in view.batches
                ]
        finally:
            conn.close()

        with tempfile.TemporaryDirectory() as tmpdir:
            path = build_template(
                Path(tmpdir) / "delivery_schedule_template.xlsx",
                plates=kp_plates,
                batches=batches,
            )
            return path.read_bytes()

    def import_draft(
        self, kp_id: int, file_bytes: bytes, *, user: dict
    ) -> ImportDraftResponse:
        """Разбор XLSX → черновик партий; в БД ничего не пишет."""
        if not file_bytes:
            raise DeliveryScheduleValidationError("Пустой файл импорта")

        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            offer = self._fetch_offer(cur, kp_id)
            if offer is None:
                raise DeliveryScheduleNotFoundError(f"КП №{kp_id} не найдено")
            assert_offer_read_access(user, offer)
            kp_plates = self._load_kp_plates_for_import(cur, kp_id)
        finally:
            conn.close()

        try:
            drafts, unmatched = parse_template(file_bytes, kp_plates)
        except RuntimeError:
            raise
        except Exception as exc:
            raise DeliveryScheduleValidationError(
                "Не удалось разобрать XLSX-шаблон. Проверьте формат файла."
            ) from exc

        return ImportDraftResponse(
            batches=[
                BatchDraftOut(
                    name=draft.name,
                    deliver_from=draft.deliver_from,
                    deliver_to=draft.deliver_to,
                    produce_by=draft.produce_by,
                    items=[
                        BatchDraftItemOut(
                            plate_id=item.plate_id,
                            plate_name=item.plate_name,
                            qty=item.qty,
                        )
                        for item in draft.items
                    ],
                )
                for draft in drafts
            ],
            unmatched_rows=[
                UnmatchedRowOut(
                    row_number=row.row_number,
                    reason=row.reason,
                    raw=row.raw,
                )
                for row in unmatched
            ],
        )

    def generate_document(
        self,
        kp_id: int,
        fmt: DocumentFmt,
        user: dict,
    ) -> Path:
        """Генерирует XLSX/PDF графика в ``outputs_dir``; существующие не затирает."""
        if fmt not in ("xlsx", "pdf"):
            raise DeliveryScheduleValidationError(
                f"Неподдерживаемый формат документа: {fmt}"
            )

        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            offer = self._fetch_offer(cur, kp_id)
            if offer is None:
                raise DeliveryScheduleNotFoundError(f"КП №{kp_id} не найдено")
            assert_offer_read_access(user, offer)

            schedule = self._fetch_schedule(cur, kp_id)
            if schedule is None:
                raise DeliveryScheduleNotFoundError(
                    f"График поставки для КП №{kp_id} не найден"
                )
            view = self._build_view(cur, schedule)
            customer_name = (offer.get("customer_name") or None)
        finally:
            conn.close()

        payload = view.model_dump()
        if customer_name:
            payload["customer_name"] = str(customer_name).strip() or None

        target = self._unique_document_path(kp_id, fmt)
        if fmt == "xlsx":
            return build_xlsx_document(payload, target)
        return build_pdf_document(payload, target)

    def _unique_document_path(self, kp_id: int, ext: str) -> Path:
        """Имя ``График_КП{id}_ред_YYYY-MM-DD.{ext}``; при коллизии — ``_HHmmss``."""
        today = date.today().isoformat()
        stem = f"График_КП{kp_id}_ред_{today}"
        candidate = self.outputs_dir / f"{stem}.{ext}"
        if not candidate.exists():
            return candidate

        stamp = datetime.now().strftime("%H%M%S")
        candidate = self.outputs_dir / f"{stem}_{stamp}.{ext}"
        if not candidate.exists():
            return candidate

        n = 2
        while True:
            candidate = self.outputs_dir / f"{stem}_{stamp}_{n}.{ext}"
            if not candidate.exists():
                return candidate
            n += 1

    # ------------------------------------------------------------------
    # traffic light
    # ------------------------------------------------------------------

    def _resolve_today(self, today: str | None) -> str:
        if today is not None:
            return date.fromisoformat(today).isoformat()
        if self._today_override is not None:
            return date.fromisoformat(self._today_override).isoformat()
        return date.today().isoformat()

    def _enrich_with_traffic_light(
        self,
        view: DeliveryScheduleView,
        kp_id: int,
        *,
        today: str | None = None,
    ) -> DeliveryScheduleView:
        """Считает светофор; при сбое источников — statuses null + degraded flag."""
        if not view.batches:
            return view

        try:
            today_iso = self._resolve_today(today)
            plates = self._load_plates_meta(kp_id)
            produced = self._load_produced_by_plate_id(kp_id, plates)
            occupancy = self._load_occupancy()
            workdays = self._collect_workdays(
                today_iso=today_iso,
                produce_by_dates=[b.produce_by for b in view.batches],
            )

            batch_inputs: list[BatchInput] = []
            item_flags: list[list[bool]] = []

            for batch in view.batches:
                items_in: list[BatchItemInput] = []
                flags: list[bool] = []
                for item in batch.items:
                    meta = plates.get(int(item.plate_id))
                    plate_qty = int(meta["qty"]) if meta is not None else 0
                    length_m = float(meta["length_m"] or 0.0) if meta is not None else 0.0
                    # R4: позиция КП уменьшилась — не падаем, помечаем changed.
                    changed = int(item.qty) > plate_qty
                    flags.append(changed)
                    check_qty = min(int(item.qty), plate_qty) if plate_qty >= 0 else 0
                    items_in.append(
                        BatchItemInput(
                            plate_id=int(item.plate_id),
                            qty=max(0, check_qty),
                            length_m=length_m,
                        )
                    )
                item_flags.append(flags)
                batch_inputs.append(
                    BatchInput(
                        id=int(batch.id),
                        name=batch.name,
                        produce_by=batch.produce_by,
                        items=items_in,
                    )
                )

            checks = check_batches(
                batches=batch_inputs,
                occupancy=occupancy,
                workdays=workdays,
                produced=produced,
                today=today_iso,
            )
        except _TRAFFIC_LIGHT_SOURCE_ERRORS:
            logger.exception(
                "Светофор графика поставки недоступен для КП %s — отдаём без status",
                kp_id,
            )
            cleared = [
                batch.model_copy(
                    update={"status": None, "ready_date": None, "hint": None}
                )
                for batch in view.batches
            ]
            return view.model_copy(
                update={"batches": cleared, "traffic_light_degraded": True}
            )

        enriched: list[BatchOut] = []
        for batch, check, flags in zip(view.batches, checks, item_flags, strict=True):
            new_items = [
                item.model_copy(update={"changed": flags[idx]})
                for idx, item in enumerate(batch.items)
            ]
            batch_changed = any(flags)
            enriched.append(
                batch.model_copy(
                    update={
                        "items": new_items,
                        "status": check.status,
                        "ready_date": check.ready_date,
                        "hint": check.hint,
                        "changed": batch_changed,
                    }
                )
            )
        return view.model_copy(
            update={"batches": enriched, "traffic_light_degraded": False}
        )

    def _load_plates_meta(self, kp_id: int) -> dict[int, dict[str, Any]]:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute(
                """
                SELECT id, plate_name, qty, length_m, width_m, load_class
                FROM kp_plates
                WHERE kp_id = ?
                """,
                (int(kp_id),),
            )
            result: dict[int, dict[str, Any]] = {}
            for row in cur.fetchall():
                result[int(row["id"])] = {
                    "plate_name": row["plate_name"],
                    "qty": int(row["qty"] or 0),
                    "length_m": row["length_m"],
                    "width_m": row["width_m"],
                    "load_class": row["load_class"],
                }
            return result
        finally:
            conn.close()

    def _load_produced_by_plate_id(
        self,
        kp_id: int,
        plates: dict[int, dict[str, Any]],
    ) -> dict[int, int]:
        """produced[plate_id]: on_sgp по identity, распределённый пропорционально qty.

        Сумма allocated по группе = min(on_sgp, Σ qty); без дублирования полного
        on_sgp на каждый plate_id одной марки.
        """
        try:
            positions = KpReadinessService(db_path=self.db_path).list_positions(kp_id)
        except _TRAFFIC_LIGHT_SOURCE_ERRORS:
            raise
        except Exception as exc:
            raise RuntimeError(
                f"KpReadinessService.list_positions недоступен для КП {kp_id}"
            ) from exc

        by_identity: dict[tuple[Any, ...], int] = {}
        for pos in positions:
            key = self._plate_identity_key(
                pos.plate_name, pos.length_m, pos.width_m, pos.load_class
            )
            by_identity[key] = int(pos.on_sgp or 0)

        groups: dict[tuple[Any, ...], list[tuple[int, int]]] = {}
        for plate_id, meta in plates.items():
            key = self._plate_identity_key(
                meta.get("plate_name"),
                meta.get("length_m"),
                meta.get("width_m"),
                meta.get("load_class"),
            )
            qty = int(meta.get("qty") or 0)
            groups.setdefault(key, []).append((int(plate_id), qty))

        produced: dict[int, int] = {}
        for key, members in groups.items():
            on_sgp = by_identity.get(key, 0)
            qtys = [q for _, q in members]
            allocated = self._allocate_largest_remainder(on_sgp, qtys)
            for (plate_id, _), alloc in zip(members, allocated, strict=True):
                produced[plate_id] = alloc
        return produced

    @staticmethod
    def _allocate_largest_remainder(total: int, weights: list[int]) -> list[int]:
        """Пропорциональное целое распределение: Σ result = min(total, Σ weights)."""
        n = len(weights)
        if n == 0:
            return []
        sum_w = sum(weights)
        if total <= 0 or sum_w <= 0:
            return [0] * n
        capped = min(int(total), sum_w)
        exact = [capped * w / sum_w for w in weights]
        floors = [int(x) for x in exact]
        leftover = capped - sum(floors)
        order = sorted(
            range(n),
            key=lambda i: (-(exact[i] - floors[i]), i),
        )
        for i in order[:leftover]:
            floors[i] += 1
        return floors

    @staticmethod
    def _plate_identity_key(
        plate_name: Any,
        length_m: Any,
        width_m: Any,
        load_class: Any,
    ) -> tuple[Any, ...]:
        try:
            length_key = round(float(length_m or 0), 3)
        except (TypeError, ValueError):
            length_key = 0.0
        try:
            width_key = round(float(width_m or 0), 3)
        except (TypeError, ValueError):
            width_key = 0.0
        return (str(plate_name or ""), length_key, width_key, load_class)

    @staticmethod
    def _load_occupancy() -> dict[str, dict]:
        """days_info из production-календаря (max с day_capacity overrides)."""
        try:
            calendar = PlanDistributionService().get_global_calendar_info(
                PlanRepository()
            )
        except _TRAFFIC_LIGHT_SOURCE_ERRORS:
            raise
        except Exception as exc:
            raise RuntimeError("get_global_calendar_info недоступен") from exc
        if not calendar:
            return {}
        days_info = calendar.get("days_info") or {}
        if not isinstance(days_info, dict):
            return {}
        return days_info

    @staticmethod
    def _collect_workdays(
        *,
        today_iso: str,
        produce_by_dates: list[str],
    ) -> set[str]:
        today_d = date.fromisoformat(today_iso)
        end = today_d + timedelta(days=_WORKDAYS_HORIZON_DAYS)
        for raw in produce_by_dates:
            try:
                pb = date.fromisoformat(raw)
            except ValueError:
                continue
            end = max(end, pb + timedelta(days=_WORKDAYS_AFTER_PRODUCE_BY))

        holidays = load_holidays()
        extra_workdays = load_extra_workdays()
        workdays: set[str] = set()
        current = today_d
        while current <= end:
            if is_working_day(current, holidays, extra_workdays):
                workdays.add(current.isoformat())
            current += timedelta(days=1)
        return workdays

    # ------------------------------------------------------------------
    # validation / DB helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _fetch_offer(cur: sqlite3.Cursor, kp_id: int) -> dict[str, Any] | None:
        cur.execute(
            """
            SELECT
                ko.kp_id,
                ko.customer_name,
                ko.manager_name,
                m.status,
                m.owner_user_id
            FROM KP_offers ko
            LEFT JOIN kp_meta m ON ko.kp_id = m.kp_id
            WHERE ko.kp_id = ?
            """,
            (int(kp_id),),
        )
        row = cur.fetchone()
        return dict(row) if row else None

    @staticmethod
    def _fetch_schedule(cur: sqlite3.Cursor, kp_id: int) -> sqlite3.Row | None:
        cur.execute(
            """
            SELECT id, kp_id, invoice_number, contract_number, status,
                   created_at, updated_at
            FROM delivery_schedule
            WHERE kp_id = ?
            """,
            (int(kp_id),),
        )
        return cur.fetchone()

    @staticmethod
    def _load_plate_qty_map(cur: sqlite3.Cursor, kp_id: int) -> dict[int, dict[str, Any]]:
        cur.execute(
            """
            SELECT id, plate_name, qty
            FROM kp_plates
            WHERE kp_id = ?
            """,
            (int(kp_id),),
        )
        result: dict[int, dict[str, Any]] = {}
        for row in cur.fetchall():
            result[int(row["id"])] = {
                "plate_name": row["plate_name"],
                "qty": int(row["qty"]),
            }
        return result

    @staticmethod
    def _load_kp_plates_for_import(
        cur: sqlite3.Cursor, kp_id: int
    ) -> list[dict[str, Any]]:
        """Список позиций КП для ``parse_template`` (ключ ``id`` + ``plate_name``)."""
        cur.execute(
            """
            SELECT id, plate_name, qty, length_m, width_m, load_class
            FROM kp_plates
            WHERE kp_id = ?
            ORDER BY id ASC
            """,
            (int(kp_id),),
        )
        return [dict(row) for row in cur.fetchall()]

    @staticmethod
    def _validate_batches_against_plates(
        payload: DeliverySchedulePut,
        plate_qty: dict[int, dict[str, Any]],
    ) -> None:
        totals: dict[int, int] = {}
        for batch in payload.batches:
            for item in batch.items:
                plate_id = int(item.plate_id)
                if plate_id not in plate_qty:
                    raise DeliveryScheduleValidationError(
                        f"Позиция plate_id={plate_id} не принадлежит этому КП"
                    )
                totals[plate_id] = totals.get(plate_id, 0) + int(item.qty)

        for plate_id, total in totals.items():
            allowed = int(plate_qty[plate_id]["qty"])
            if total > allowed:
                name = plate_qty[plate_id]["plate_name"] or f"#{plate_id}"
                raise DeliveryScheduleValidationError(
                    f"Сумма по позиции «{name}» (plate_id={plate_id}): "
                    f"{total} шт. превышает количество в КП ({allowed} шт.)"
                )

    @staticmethod
    def _upsert_schedule_header(
        cur: sqlite3.Cursor,
        kp_id: int,
        payload: DeliverySchedulePut,
        now: str,
    ) -> int:
        cur.execute(
            "SELECT id FROM delivery_schedule WHERE kp_id = ?",
            (int(kp_id),),
        )
        row = cur.fetchone()
        invoice = payload.invoice_number
        contract = payload.contract_number
        if row is not None:
            schedule_id = int(row["id"])
            cur.execute(
                """
                UPDATE delivery_schedule
                SET invoice_number = ?,
                    contract_number = ?,
                    updated_at = ?
                WHERE id = ?
                """,
                (invoice, contract, now, schedule_id),
            )
            return schedule_id

        cur.execute(
            """
            INSERT INTO delivery_schedule (
                kp_id, invoice_number, contract_number, status,
                created_at, updated_at
            ) VALUES (?, ?, ?, 'draft', ?, ?)
            """,
            (int(kp_id), invoice, contract, now, now),
        )
        return int(cur.lastrowid)

    @staticmethod
    def _replace_batches(
        cur: sqlite3.Cursor,
        schedule_id: int,
        payload: DeliverySchedulePut,
    ) -> None:
        cur.execute(
            "DELETE FROM delivery_batch WHERE schedule_id = ?",
            (int(schedule_id),),
        )
        for index, batch in enumerate(payload.batches):
            sort_order = int(batch.sort_order) if batch.sort_order is not None else index
            cur.execute(
                """
                INSERT INTO delivery_batch (
                    schedule_id, name, deliver_from, deliver_to,
                    produce_by, sort_order
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    int(schedule_id),
                    batch.name,
                    batch.deliver_from,
                    batch.deliver_to,
                    batch.produce_by,
                    sort_order,
                ),
            )
            batch_id = int(cur.lastrowid)
            for item in batch.items:
                cur.execute(
                    """
                    INSERT INTO delivery_batch_item (batch_id, plate_id, qty)
                    VALUES (?, ?, ?)
                    """,
                    (batch_id, int(item.plate_id), int(item.qty)),
                )

    def _build_view(
        self,
        cur: sqlite3.Cursor,
        schedule: sqlite3.Row,
    ) -> DeliveryScheduleView:
        schedule_id = int(schedule["id"])
        cur.execute(
            """
            SELECT id, name, deliver_from, deliver_to, produce_by, sort_order
            FROM delivery_batch
            WHERE schedule_id = ?
            ORDER BY sort_order ASC, id ASC
            """,
            (schedule_id,),
        )
        batch_rows = cur.fetchall()
        batches: list[BatchOut] = []
        for brow in batch_rows:
            batch_id = int(brow["id"])
            cur.execute(
                """
                SELECT
                    i.plate_id,
                    i.qty,
                    p.plate_name
                FROM delivery_batch_item i
                LEFT JOIN kp_plates p ON p.id = i.plate_id
                WHERE i.batch_id = ?
                ORDER BY i.id ASC
                """,
                (batch_id,),
            )
            items = [
                BatchItemOut(
                    plate_id=int(irow["plate_id"]),
                    qty=int(irow["qty"]),
                    plate_name=irow["plate_name"],
                    changed=False,
                )
                for irow in cur.fetchall()
            ]
            batches.append(
                BatchOut(
                    id=batch_id,
                    name=str(brow["name"]),
                    deliver_from=str(brow["deliver_from"]),
                    deliver_to=str(brow["deliver_to"]),
                    produce_by=str(brow["produce_by"]),
                    items=items,
                    sort_order=int(brow["sort_order"] or 0),
                    status=None,
                    ready_date=None,
                    hint=None,
                    changed=False,
                )
            )

        return DeliveryScheduleView(
            id=schedule_id,
            kp_id=int(schedule["kp_id"]),
            invoice_number=schedule["invoice_number"],
            contract_number=schedule["contract_number"],
            status=schedule["status"] or "draft",
            batches=batches,
            updated_at=str(schedule["updated_at"]),
        )
