from __future__ import annotations

import asyncio
import json
import logging
import os
import shutil
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Literal

from app.core.settings import get_settings
from app.domain.models.plate_order import PlateOrder as AppPlateOrder
from app.repositories.kp_archive_repository import ArchiveSection, KpArchiveRepository
from app.schemas.archive import (
    ArchiveBridgePileItem,
    ArchiveFbsItem,
    ArchiveFileKind,
    ArchiveOfferDetails,
    ArchiveOfferFinance,
    ArchiveOfferListItem,
    ArchiveMarchItem,
    ArchivePileItem,
    ArchivePlateItem,
    ArchiveStepItem,
    ArchiveProductTypeFilter,
    ArchiveSearchResponse,
    CapacityDayInfo,
    CapacitySnapshotResponse,
    KpReadinessPositionsResponse,
    SpecificationChoiceIn,
    SpecificationChoiceOut,
    SpecificationView,
)
from app.repositories.plan_repository import PlanRepository
from app.services.capacity_gate_service import (
    CapacitySnapshot,
    build_capacity_snapshot,
)
from app.services.plan_distribution_service import PlanDistributionService
from app.security.offer_access import (
    assert_offer_read_access,
    assert_offer_write_access,
    list_filters_for_user,
)
from app.services.file_generation_service import FileGenerationService
from app.services.optimization_service import OptimizationService
from app.services.promise_service import PromiseGateError, PromiseService
from app.repositories.supply_contract_repository import SupplyContractRepository
from app.services.counterparties_service import (
    CounterpartiesService,
    CounterpartyValidationError,
)
from core.guid_gate import InvoiceGuidBlockError, OrderLine, assert_invoice_guids
from core.embed_delivery_in_unit_price import allocate_line_delivery_kopecks
from core.invoice_export import (
    INVOICE_WAREHOUSES,
    InvoiceAction,
    _line_payable,
    build_invoice_document,
    card_offer_from_kp,
    invoice_snapshot_hash,
)
from core.specification_buyer import buyer_preamble_for_print
from core.specification_text import (
    SpecificationChoice,
    SpecificationTextError,
    choice_from_document,
    render_specification,
    specification_document,
)
from core.specification_xlsx import (
    SpecificationHeader,
    SpecificationPayableLine,
    build_specification_xlsx,
)
from core.supply_contract import attach_contract_number

STATUS_ARCHIVE = "в архиве"
STATUS_ON_APPROVAL = "на согласовании"
MSG_EXPORT_ONLY_ARCHIVE = "Отправить в 1С можно только из статуса «в архиве»"
MSG_ORDER_IDS_MISSING = "Исправление не отправляется: номер заказа из 1С ещё не записан"
MSG_NEED_COUNTERPARTY = "Сначала занесите контрагента"
MSG_NEED_CONTRACT = "Нет действующего договора поставки"
MSG_SPEC_STATUS = "Сохранить спецификацию можно только в статусе «на согласовании»"
MSG_SPEC_PICKUP = (
    "Самовывоз нельзя сохранить: уберите доставку в конструкторе или выберите доставку на объект."
)
MSG_SPEC_REQUIRED = "Сначала сохраните спецификацию"
MSG_SPEC_STALE = (
    "Условия оплаты в спецификации устарели: откройте спецификацию и сохраните их снова"
)
from core.production.promise_buckets import OccupancyUnavailableError
from core.delivery_schedule_check import BatchItemInput
from core.plate_order_context import PlateOrderContext, run_in_order_context
from core.ports.visualization import get_visualize_plan
from core.execution_terms import parse_execution_terms
from core.commercial_offer import generate_commercial_offer_pdf
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.cargo_delivery_pricing import delivery_service_charge_rub, total_order_cargo_weight_kg
from core.kp_order_data import order_data_from_kp_info
from core.gantt_excel import create_gantt_excel
from core.production_capacity import MAX_TRACK_LENGTH_M, TRACKS_PER_DAY_DEFAULT
from core.work_calendar import is_working_day, load_extra_workdays, load_holidays


logger = logging.getLogger(__name__)


class ArchiveError(Exception):
    """Базовое исключение сервиса архива."""


class ArchiveNotFoundError(ArchiveError):
    """КП не найдено в БД."""


class ArchiveValidationError(ArchiveError):
    """Ошибка валидации бизнес-правил (например, недопустимый статус)."""


@dataclass(frozen=True)
class SpecificationDownload:
    content: bytes
    filename: str


class ArchiveService:
    """Бизнес-логика раздела «Архив» (аналог bot/handlers/archive.py)."""

    def __init__(
        self,
        repository: KpArchiveRepository | None = None,
        outputs_dir: Path | None = None,
        optimization_service: OptimizationService | None = None,
        file_generation_service: FileGenerationService | None = None,
        promise_service: PromiseService | None = None,
        export_dir: Path | None = None,
    ) -> None:
        settings = get_settings()
        self.repository = repository or KpArchiveRepository()
        self.export_dir = export_dir
        self.outputs_dir = outputs_dir or settings.outputs_dir
        self.outputs_dir.mkdir(parents=True, exist_ok=True)
        self.optimization_service = optimization_service or OptimizationService()
        self.file_generation_service = file_generation_service or FileGenerationService()
        self._promise_service = promise_service
        # ISO YYYY-MM-DD для детерминированных тестов гейта ёмкости.
        self._today_override: str | None = None

    # ---------- Списки и карточка ----------

    def list_offers(
        self,
        section: ArchiveSection,
        *,
        user: dict,
        product_type: ArchiveProductTypeFilter = "all",
    ) -> list[ArchiveOfferListItem]:
        list_filters = list_filters_for_user(user)
        raw_items = self.repository.list_by_section(
            section,
            product_type=product_type,
            **list_filters,
        )
        return [self._to_list_item(raw) for raw in raw_items]

    def get_details(self, kp_id: int, *, user: dict) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_read_access(user, raw)
        return self._to_details(raw)

    def resume_as_draft(self, kp_id: int, *, user: dict) -> dict:
        """Hydrate a commercial draft from «в архиве» or «на согласовании»."""
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)

        from app.services.commercial_workflow_service import CommercialWorkflowService

        try:
            return CommercialWorkflowService().hydrate_draft_from_saved_kp(
                kp_id,
                owner_user_id=int(user["id"]),
            )
        except ValueError as exc:
            raise ArchiveValidationError(str(exc)) from exc

    def get_readiness_positions(self, kp_id: int, *, user: dict) -> KpReadinessPositionsResponse:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_read_access(user, raw)
        status = raw.get("status") or ""
        from app.services.kp_readiness_service import KpReadinessService

        items = KpReadinessService(db_path=self.repository.db_path).list_positions(
            kp_id, status=status
        )
        return KpReadinessPositionsResponse(items=items, count=len(items))

    def search(
        self,
        *,
        user: dict,
        kp_id: int | None = None,
        customer: str | None = None,
    ) -> ArchiveSearchResponse:
        list_filters = list_filters_for_user(user)
        if kp_id is not None:
            raw = self.repository.get_by_id(kp_id)
            if not raw:
                rows: list[dict] = []
            else:
                assert_offer_read_access(user, raw)
                rows = [raw]
            return ArchiveSearchResponse(
                mode="number",
                items=[self._to_list_item(r) for r in rows],
                total=len(rows),
                truncated=False,
            )

        name = (customer or "").strip()
        rows, total = self.repository.search_by_customer_name(name, limit=50, **list_filters)
        return ArchiveSearchResponse(
            mode="customer",
            items=[self._to_list_item(raw) for raw in rows],
            total=total,
            truncated=total > 50,
        )

    # ---------- Мутации ----------

    def update_discount(self, kp_id: int, discount: float, *, user: dict) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        if not 0 <= discount <= 100:
            raise ArchiveValidationError("Процент скидки должен быть от 0 до 100")
        if not self.repository.update_discount(kp_id, discount):
            raise ArchiveNotFoundError(
                f"Не удалось обновить скидку. КП №{kp_id} не найдено или пустое."
            )
        return self.get_details(kp_id, user=user)

    def update_logistics_cost(
        self,
        kp_id: int,
        logistics_cost: float,
        *,
        user: dict,
        pile_logistics_cost: float | None = None,
        pile_trip_overrides: dict | None = None,
        long_pile_delivery: dict | None = None,
    ) -> ArchiveOfferDetails:
        """Обновляет тарифы рейсов и суммы заказа (без авто-PDF)."""
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        trip = max(0.0, float(logistics_cost or 0.0))
        logistics_kwargs: dict = {
            "pile_logistics_cost": pile_logistics_cost,
            "pile_trip_overrides": pile_trip_overrides,
        }
        if long_pile_delivery is not None:
            logistics_kwargs["long_pile_delivery"] = long_pile_delivery
        if not self.repository.update_logistics_cost(
            kp_id,
            trip,
            **logistics_kwargs,
        ):
            raise ArchiveNotFoundError(
                f"Не удалось обновить стоимость рейса. КП №{kp_id} не найдено или пустое."
            )
        return self.get_details(kp_id, user=user)

    def bind_counterparty(
        self, kp_id: int, counterparty_id: int, *, user: dict
    ) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        if raw.get("status") != "в архиве":
            raise ArchiveValidationError(
                "Привязать контрагента можно только у КП в статусе «в архиве»"
            )
        try:
            row = CounterpartiesService(
                db_path=self.repository.db_path
            ).require_active_client(counterparty_id)
        except CounterpartyValidationError as exc:
            raise ArchiveValidationError(str(exc)) from exc
        if not self.repository.update_counterparty(
            kp_id,
            counterparty_id=int(row["id"]),
            customer_name=str(row["name"]),
            customer_inn=row.get("inn"),
            customer_kpp=row.get("kpp"),
        ):
            raise ArchiveNotFoundError(
                f"Не удалось привязать контрагента. КП №{kp_id} не найдено."
            )
        return self.get_details(kp_id, user=user)

    def create_and_bind_counterparty(
        self,
        kp_id: int,
        *,
        name: str,
        code_1c: str,
        inn: str | None = None,
        kpp: str | None = None,
        user: dict,
    ) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        if raw.get("status") != "в архиве":
            raise ArchiveValidationError(
                "Привязать контрагента можно только у КП в статусе «в архиве»"
            )
        try:
            row = CounterpartiesService(
                db_path=self.repository.db_path
            ).resolve_active_client_for_bind(
                name=name, code_1c=code_1c, inn=inn, kpp=kpp
            )
        except CounterpartyValidationError as exc:
            raise ArchiveValidationError(str(exc)) from exc
        except ValueError as exc:
            raise ArchiveValidationError(str(exc)) from exc
        if not self.repository.update_counterparty(
            kp_id,
            counterparty_id=int(row["id"]),
            customer_name=str(row["name"]),
            customer_inn=row.get("inn"),
            customer_kpp=row.get("kpp"),
        ):
            raise ArchiveNotFoundError(
                f"Не удалось привязать контрагента. КП №{kp_id} не найдено."
            )
        return self.get_details(kp_id, user=user)

    def delete_offer(self, kp_id: int, *, user: dict) -> None:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} уже удалено или не существует")
        assert_offer_write_access(user, raw)
        self._promises().release_on_delete(kp_id)
        if not self.repository.delete(kp_id):
            raise ArchiveNotFoundError(f"КП №{kp_id} уже удалено или не существует")

    def export_invoice(
        self,
        kp_id: int,
        *,
        user: dict,
        warehouse: str | None = None,
    ) -> ArchiveOfferDetails:
        raw = self._offer_for_invoice_write(kp_id, user)
        if raw.get("status") != STATUS_ARCHIVE:
            raise ArchiveValidationError(MSG_EXPORT_ONLY_ARCHIVE)
        offer, contract = self._invoice_offer(raw)
        name = _require_invoice_warehouse(warehouse)
        document, digest = self._build_stamped_invoice(
            raw, offer, contract, action="create", warehouse=name
        )
        self._write_invoice_file(kp_id, "create", document)
        try:
            self.repository.commit_invoice_snapshot(
                kp_id, digest, status=STATUS_ON_APPROVAL, warehouse=name
            )
        except Exception as exc:
            logger.exception("invoice status commit failed for kp_id=%s", kp_id)
            raise ArchiveError(f"Не удалось обновить статус КП №{kp_id}") from exc
        return self.get_details(kp_id, user=user)

    def export_invoice_correction(
        self,
        kp_id: int,
        *,
        user: dict,
        warehouse: str | None = None,
    ) -> ArchiveOfferDetails:
        raw = self._offer_for_invoice_write(kp_id, user)
        if raw.get("status") != STATUS_ON_APPROVAL:
            raise ArchiveValidationError(
                "Исправление можно отправить только из статуса «на согласовании»"
            )
        ids = _invoice_order_ids(raw)
        if ids is None:
            raise ArchiveValidationError(MSG_ORDER_IDS_MISSING)
        number, uid_order, uid_kp = ids
        current = invoice_snapshot_hash(card_offer_from_kp(raw))
        if str(raw.get("invoice_snapshot_hash") or "") == current:
            raise ArchiveValidationError("Состав не изменился")
        offer, contract = self._invoice_offer(raw)
        name = _warehouse_for_correction(raw, warehouse)
        document, digest = self._build_stamped_invoice(
            raw,
            offer,
            contract,
            action="update",
            warehouse=name,
            order_number=number,
            uid_order=uid_order,
            uid_kp=uid_kp,
        )
        self._write_invoice_file(kp_id, "update", document)
        try:
            self.repository.commit_invoice_snapshot(kp_id, digest, warehouse=name)
        except Exception as exc:
            logger.exception("invoice correction commit failed for kp_id=%s", kp_id)
            raise ArchiveError(f"Не удалось обновить снимок КП №{kp_id}") from exc
        return self.get_details(kp_id, user=user)

    def _build_stamped_invoice(
        self,
        raw: dict,
        offer: dict,
        contract: dict,
        *,
        action: InvoiceAction,
        warehouse: str,
        order_number: str | None = None,
        uid_order: str | None = None,
        uid_kp: str | None = None,
    ) -> tuple[dict[str, Any], str]:
        shares = self._invoice_delivery_kopecks(raw, offer)
        document, digest = build_invoice_document(
            offer,
            action=action,
            order_number=order_number,
            uid_order=uid_order,
            uid_kp=uid_kp,
            warehouse=warehouse,
            line_delivery_kopecks=shares,
        )
        stamped = attach_contract_number(
            document,
            int(raw["counterparty_id"]),
            contract,
        )
        return stamped, digest

    def _invoice_delivery_kopecks(self, raw: dict, offer: dict) -> list[int]:
        totals = self._calculate_delivery_totals(raw)
        reason = _unready_delivery_reason(totals)
        if reason:
            raise ArchiveValidationError(reason)
        long_enabled = bool(totals.get("long_pile_delivery_enabled"))
        lookup = self._pile_catalog_lookup() if long_enabled else None
        qtys, types, length_keys = _allocation_rows(
            raw, long_enabled=long_enabled, catalog_lookup=lookup
        )
        if len(qtys) != len(offer.get("lines") or []):
            raise ArchiveValidationError("Не удалось разложить доставку по строкам счёта")
        long_amounts = _ready_long_pile_amounts(totals)
        kopecks = allocate_line_delivery_kopecks(
            qty_by_index=qtys,
            product_type_by_index=types,
            plate_delivery_total=float(totals.get("plate_delivery_total") or 0.0),
            pile_delivery_total=float(totals.get("pile_delivery_total") or 0.0),
            fbs_delivery_total=float(totals.get("fbs_lm_delivery_total") or 0.0),
            long_pile_delivery_by_length=long_amounts,
            length_key_by_index=length_keys,
        )
        if sum(kopecks) != _expected_delivery_kopecks(totals, long_amounts):
            raise ArchiveValidationError("Не удалось разложить доставку по строкам счёта")
        return kopecks

    def _calculate_delivery_totals(self, raw: dict) -> dict[str, Any]:
        from core.commercial_pricing import (
            calculate_total_cost,
            coerce_fbs_lm_delivery_enabled,
            coerce_long_pile_delivery_enabled,
        )
        from core.pile_trip_pricing import coerce_pile_trip_overrides

        kp_id = raw.get("kp_id")
        return calculate_total_cost(
            order_data_from_kp_info(raw),
            float(raw.get("discount_percent") or 0.0),
            logistics_cost=max(0.0, float(raw.get("logistics_cost") or 0.0)),
            db_path=self.repository.db_path,
            require_all_priced=False,
            pile_logistics_cost=max(0.0, float(raw.get("pile_logistics_cost") or 0.0)),
            pile_trip_overrides=coerce_pile_trip_overrides(raw.get("pile_trip_overrides_json")),
            pile_catalog_db_path=self.repository.db_path,
            fbs_lm_delivery_enabled=coerce_fbs_lm_delivery_enabled(
                raw.get("fbs_lm_delivery_enabled")
            ),
            weight_catalog_db_path=self.repository.db_path,
            long_pile_delivery_enabled=coerce_long_pile_delivery_enabled(
                raw.get("long_pile_delivery_enabled")
            ),
            long_pile_delivery=raw.get("long_pile_delivery_json"),
            kp_id=int(kp_id) if kp_id is not None else None,
        )

    def _pile_catalog_lookup(self):
        from core.pile_catalog import load_pile_catalog, resolve_catalog_for_mark

        entries = load_pile_catalog(self.repository.db_path)
        return lambda mark: resolve_catalog_for_mark(mark, entries)

    def _offer_for_invoice_write(self, kp_id: int, user: dict) -> dict:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        return raw

    def _invoice_offer(self, raw: dict) -> tuple[dict[str, Any], dict]:
        counterparty_id = raw.get("counterparty_id")
        if counterparty_id is None:
            raise ArchiveValidationError(MSG_NEED_COUNTERPARTY)
        try:
            party = CounterpartiesService(
                db_path=self.repository.db_path
            ).require_active_client(counterparty_id)
        except CounterpartyValidationError as exc:
            raise ArchiveValidationError(str(exc)) from exc
        contract = SupplyContractRepository(
            db_path=self.repository.db_path
        ).get_active_by_counterparty(int(counterparty_id))
        if contract is None:
            raise ArchiveValidationError(MSG_NEED_CONTRACT)
        offer = card_offer_from_kp(raw)
        offer["counterparty"] = {
            "name": party.get("name") or raw.get("customer_name"),
            "code": party.get("code_1c"),
            "guid": party.get("guid_1c"),
            "inn": party.get("inn") or raw.get("customer_inn"),
            "kpp": party.get("kpp") or raw.get("customer_kpp"),
        }
        order_lines = [
            OrderLine(
                str(line.get("product_kind") or "plate"),
                str(line.get("name") or ""),
                int(line.get("qty") or 1),
            )
            for line in offer["lines"]
        ]
        conn = self._open_guid_db()
        try:
            report = assert_invoice_guids(order_lines, conn)
        except InvoiceGuidBlockError as exc:
            raise ArchiveValidationError(str(exc)) from exc
        finally:
            conn.close()
        for line, ready in zip(offer["lines"], report.ready):
            line["guid"] = ready.guid
        return offer, contract

    def _write_invoice_file(self, kp_id: int, action: str, document: dict) -> Path:
        settings = get_settings()
        directory = self.export_dir if self.export_dir is not None else Path(settings.exchange_export_dir)
        prefix = "new" if action == "create" else "delta"
        try:
            directory.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%dT%H%M%S%f")
            path = directory / f"{prefix}_{kp_id}_{stamp}.json"
            path.write_text(
                json.dumps(document, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except OSError as exc:
            raise ArchiveError("Не удалось записать файл счёта") from exc
        return path

    def _correction_pending(self, raw: dict) -> bool:
        if raw.get("status") != STATUS_ON_APPROVAL:
            return False
        stored = str(raw.get("invoice_snapshot_hash") or "").strip()
        if not stored:
            return False
        return invoice_snapshot_hash(card_offer_from_kp(raw)) != stored

    def _invoice_export_block(self, raw: dict) -> str | None:
        if raw.get("status") != STATUS_ARCHIVE:
            return None
        if raw.get("counterparty_id") is None:
            return MSG_NEED_COUNTERPARTY
        try:
            contract = SupplyContractRepository(
                db_path=self.repository.db_path
            ).get_active_by_counterparty(int(raw["counterparty_id"]))
        except Exception:
            logger.exception("Не удалось прочитать договор для КП %s", raw.get("kp_id"))
            return None
        if contract is None:
            return MSG_NEED_CONTRACT
        try:
            offer = card_offer_from_kp(raw)
            order_lines = [
                OrderLine(
                    str(line.get("product_kind") or "plate"),
                    str(line.get("name") or ""),
                    int(line.get("qty") or 1),
                )
                for line in offer["lines"]
                if str(line.get("name") or "").strip()
            ]
            if not order_lines:
                return None
            conn = self._open_guid_db()
            try:
                assert_invoice_guids(order_lines, conn)
            finally:
                conn.close()
        except InvoiceGuidBlockError as exc:
            return str(exc)
        except Exception:
            logger.exception("Не удалось проверить GUID для КП %s", raw.get("kp_id"))
            return "Не удалось проверить GUID 1С"
        return None

    def _open_guid_db(self) -> sqlite3.Connection:
        """Справочник GUID — pb.db. В тестах без этого файла остаётся база КП."""
        pb_path = Path(get_settings().pb_db_path)
        if pb_path.is_file():
            return sqlite3.connect(pb_path)
        return sqlite3.connect(self.repository.db_path)

    def get_specification(self, kp_id: int, *, user: dict) -> SpecificationView:
        raw = self._specification_offer(kp_id, user, write=False)
        total, delivery_sum = self._specification_totals(raw)
        document = _load_stored_specification(raw)
        has_piles = _kp_has_piles(raw)
        grade = _uniform_concrete_grade(raw)
        if document is None:
            choice = _suggested_choice(raw, delivery_sum)
            payment, term, delivery = _suggestion_paragraphs(choice, total)
            return SpecificationView(
                saved=False,
                stale_custom=False,
                has_piles=has_piles,
                concrete_grade=grade,
                payment_paragraph=payment,
                term_paragraph=term,
                delivery_paragraph=delivery,
                choice=choice,
            )
        try:
            stored = choice_from_document(document)
            paragraphs = render_specification(
                stored, payable_total=total, has_piles=has_piles
            )
        except SpecificationTextError as exc:
            raise ArchiveValidationError(str(exc)) from exc
        current = invoice_snapshot_hash(card_offer_from_kp(raw))
        stale = stored.payment == "custom" and str(document.get("composition_hash") or "") != current
        return SpecificationView(
            saved=True,
            stale_custom=stale,
            has_piles=has_piles,
            concrete_grade=grade,
            payment_paragraph=paragraphs.payment,
            term_paragraph=paragraphs.term,
            delivery_paragraph=paragraphs.delivery,
            spec_date=str(document.get("spec_date") or "") or None,
            choice=_choice_out(document),
        )

    def save_specification(
        self,
        kp_id: int,
        payload: SpecificationChoiceIn,
        *,
        user: dict,
    ) -> SpecificationView:
        raw = self._specification_offer(kp_id, user, write=True)
        if raw.get("status") != STATUS_ON_APPROVAL:
            raise ArchiveValidationError(MSG_SPEC_STATUS)
        choice = _choice_from_payload(payload)
        has_piles = _kp_has_piles(raw)
        try:
            render_specification(choice, payable_total=0, has_piles=has_piles)
        except SpecificationTextError as exc:
            raise ArchiveValidationError(str(exc)) from exc
        _total, delivery_sum = self._specification_totals(raw)
        if choice.delivery == "pickup" and delivery_sum != 0:
            raise ArchiveValidationError(MSG_SPEC_PICKUP)
        existing = _load_stored_specification(raw)
        spec_date = str((existing or {}).get("spec_date") or "").strip()
        if not spec_date:
            spec_date = self._spec_today()
        document = specification_document(
            choice,
            spec_date=spec_date,
            composition_hash=invoice_snapshot_hash(card_offer_from_kp(raw)),
            has_piles=has_piles,
        )
        if not self.repository.set_specification_json(
            kp_id, json.dumps(document, ensure_ascii=False)
        ):
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        return self.get_specification(kp_id, user=user)

    def download_specification(self, kp_id: int, *, user: dict) -> SpecificationDownload:
        """Байты уже сохранённой спецификации. Строку не создаёт."""
        raw = self._specification_offer(kp_id, user, write=False)
        document = _load_stored_specification(raw)
        if document is None:
            raise ArchiveValidationError("Сначала сохраните спецификацию")
        total, _delivery_sum = self._specification_totals(raw)
        has_piles = _kp_has_piles(raw)
        try:
            stored = choice_from_document(document)
            paragraphs = render_specification(
                stored, payable_total=total, has_piles=has_piles
            )
            spec_date = date.fromisoformat(str(document.get("spec_date") or ""))
        except (SpecificationTextError, ValueError) as exc:
            raise ArchiveValidationError("Спецификация сохранена некорректно") from exc
        number = str(raw.get("order_number_1c") or "").strip() or None
        header = self._specification_header(
            raw,
            spec_date=spec_date,
            invoice_number=number,
            concrete_grade=_uniform_concrete_grade(raw),
        )
        content = build_specification_xlsx(
            header, self._specification_lines(raw), paragraphs
        )
        return SpecificationDownload(
            content=content,
            filename=_specification_filename(kp_id, number),
        )

    def _specification_header(
        self,
        raw: dict,
        *,
        spec_date: date,
        invoice_number: str | None,
        concrete_grade: str | None,
    ) -> SpecificationHeader:
        counterparty_id = raw.get("counterparty_id")
        if counterparty_id is None:
            raise ArchiveValidationError(MSG_NEED_COUNTERPARTY)
        contract = SupplyContractRepository(
            db_path=self.repository.db_path
        ).get_active_by_counterparty(int(counterparty_id))
        if contract is None:
            raise ArchiveValidationError(MSG_NEED_CONTRACT)
        preamble = buyer_preamble_for_print(
            legal_form=str(contract.get("legal_form") or ""),
            full_name=str(contract.get("full_name") or ""),
            signatory_name=str(contract.get("signatory_name") or ""),
            authority_basis=str(contract.get("authority_basis") or "устав"),
            signatory_position=contract.get("signatory_position"),
            signatory_verb=str(contract.get("signatory_verb") or "действующего"),
            poa_number=contract.get("poa_number"),
            poa_date=contract.get("poa_date"),
        )
        return SpecificationHeader(
            invoice_number=invoice_number,
            spec_date=spec_date,
            contract_number=str(contract.get("number") or "").strip(),
            contract_date=_parse_contract_date(contract.get("contract_date")),
            buyer_preamble=preamble,
            buyer_short_name=str(contract.get("short_name") or ""),
            buyer_inn=str(contract.get("inn") or ""),
            buyer_kpp=str(contract.get("kpp") or ""),
            buyer_signatory_position=str(contract.get("signatory_position") or ""),
            buyer_signatory_short=_short_person_name(str(contract.get("signatory_name") or "")),
            concrete_grade=concrete_grade,
        )

    def _specification_lines(self, raw: dict) -> list[SpecificationPayableLine]:
        offer = card_offer_from_kp(raw)
        shares = self._invoice_delivery_kopecks(raw, offer)
        lines = [line for line in offer.get("lines") or [] if isinstance(line, dict)]
        result: list[SpecificationPayableLine] = []
        for line, kopecks in zip(lines, shares):
            payable = _line_payable(offer, line, int(kopecks))
            qty = line.get("qty")
            count = int(qty) if isinstance(qty, int) and not isinstance(qty, bool) and qty > 0 else 0
            if count < 1:
                count = 1 if payable else 0
            price = (
                (payable / Decimal(count)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                if count
                else payable
            )
            result.append(
                SpecificationPayableLine(
                    name=str(line.get("name") or ""),
                    qty=count,
                    price=price,
                    amount=payable,
                )
            )
        return result

    def _specification_offer(self, kp_id: int, user: dict, *, write: bool) -> dict:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        if write:
            assert_offer_write_access(user, raw)
        else:
            assert_offer_read_access(user, raw)
        return raw

    def _specification_totals(self, raw: dict) -> tuple[int, int]:
        offer = card_offer_from_kp(raw)
        shares = self._invoice_delivery_kopecks(raw, offer)
        document, _digest = build_invoice_document(offer, line_delivery_kopecks=shares)
        return _rub_kopecks(document.get("СуммаДокумента")), sum(int(item) for item in shares)

    def _specification_flags(self, raw: dict) -> tuple[bool, bool]:
        document = _load_stored_specification(raw)
        if document is None:
            return False, False
        try:
            stored = choice_from_document(document)
            render_specification(
                stored, payable_total=0, has_piles=_kp_has_piles(raw)
            )
        except SpecificationTextError:
            return False, False
        if stored.payment != "custom":
            return True, False
        current = invoice_snapshot_hash(card_offer_from_kp(raw))
        return True, str(document.get("composition_hash") or "") != current

    def _spec_today(self) -> str:
        if self._today_override:
            return self._today_override
        return date.today().isoformat()

    def set_payment(self, kp_id: int, *, paid: bool, user: dict) -> ArchiveOfferDetails:
        raw = self._offer_for_invoice_write(kp_id, user)
        if raw.get("status") != STATUS_ON_APPROVAL:
            if paid:
                raise ArchiveValidationError(
                    "Отметить оплату можно только в статусе «на согласовании»"
                )
            raise ArchiveValidationError(
                "Снять оплату можно только в статусе «на согласовании»"
            )
        stamp = datetime.now().isoformat(timespec="seconds") if paid else None
        if not self.repository.set_paid_at(kp_id, stamp):
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        return self.get_details(kp_id, user=user)

    def move_to_production(self, kp_id: int, terms_input: str, *, user: dict) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        if raw.get("status") != STATUS_ON_APPROVAL:
            raise ArchiveValidationError(
                "Перевести в производство можно только из статуса «на согласовании»"
            )
        if not str(raw.get("paid_at") or "").strip():
            raise ArchiveValidationError("Сначала отметьте оплату")
        if not str(raw.get("order_number_1c") or "").strip():
            raise ArchiveValidationError("Нужен номер счёта")
        saved_spec, stale_spec = self._specification_flags(raw)
        if not saved_spec:
            raise ArchiveValidationError(MSG_SPEC_REQUIRED)
        if stale_spec:
            raise ArchiveValidationError(MSG_SPEC_STALE)
        try:
            CounterpartiesService(db_path=self.repository.db_path).require_active_client(
                raw.get("counterparty_id")
            )
        except CounterpartyValidationError as exc:
            raise ArchiveValidationError(str(exc)) from exc

        execution_terms = self._parse_execution_terms(terms_input)
        try:
            self._promises().commit_move_with_gate(
                kp_id, execution_terms, user=user, raw=raw
            )
        except (PromiseGateError, OccupancyUnavailableError) as exc:
            raise ArchiveValidationError(str(exc)) from exc
        except ArchiveError:
            raise
        except Exception as exc:
            logger.exception("move_to_production failed for kp_id=%s", kp_id)
            raise ArchiveError(
                f"Не удалось перевести КП №{kp_id} в производство"
            ) from exc

        return self.get_details(kp_id, user=user)

    def _promises(self) -> PromiseService:
        if self._promise_service is not None:
            return self._promise_service
        today = None
        if self._today_override is not None:
            today = date.fromisoformat(self._today_override)
        return PromiseService(db_path=self.repository.db_path, today=today)

    def get_capacity_snapshot(
        self,
        kp_id: int,
        *,
        user: dict,
        target: str | None = None,
    ) -> CapacitySnapshotResponse:
        """Снимок ёмкости для виджета «В производство» (старт = завтра)."""
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_read_access(user, raw)

        target_iso = self._resolve_target_iso(raw, target)
        snap = self._build_capacity_snapshot(raw, target_iso=target_iso)
        return self._snapshot_to_response(snap)

    def estimate_production(self, kp_id: int, *, user: dict) -> dict:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_read_access(user, raw)
        plates = raw.get("plates") or []
        total_length = sum((p.get("length_m") or 0) * (p.get("qty") or 1) for p in plates)
        estimated_tracks = max(1, int(round(total_length / MAX_TRACK_LENGTH_M + 0.5)))
        estimated_days = max(1, int(round(estimated_tracks / TRACKS_PER_DAY_DEFAULT + 0.5)))
        return {
            "total_length_m": total_length,
            "estimated_tracks": estimated_tracks,
            "estimated_days": estimated_days,
        }

    # ---------- Документы ----------

    async def generate_document(
        self,
        kp_id: int,
        kind: ArchiveFileKind,
        *,
        user: dict,
        plate_order_ctx: PlateOrderContext | None = None,
    ) -> Path:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_read_access(user, raw)

        order_data = order_data_from_kp_info(raw)
        if not order_data:
            raise ArchiveValidationError("В КП нет позиций для формирования документа")

        offer_number = str(kp_id)
        offer_date = raw.get("creation_date") or datetime.now().strftime("%d.%m.%Y")
        customer_name = raw.get("customer_name")
        manager_name = raw.get("manager_name")
        discount_percent = float(raw.get("discount_percent") or 0)
        logistics_cost = max(0.0, float(raw.get("logistics_cost") or 0.0))
        pile_logistics_cost = max(0.0, float(raw.get("pile_logistics_cost") or 0.0))
        from core.pile_trip_pricing import coerce_pile_trip_overrides

        pile_trip_overrides = coerce_pile_trip_overrides(raw.get("pile_trip_overrides_json"))
        from core.commercial_pricing import coerce_fbs_lm_delivery_enabled

        fbs_lm_delivery_enabled = coerce_fbs_lm_delivery_enabled(
            raw.get("fbs_lm_delivery_enabled")
        )
        from core.commercial_pricing import coerce_long_pile_delivery_enabled
        from core.pile_trip_pricing import coerce_long_pile_delivery

        long_pile_delivery_enabled = coerce_long_pile_delivery_enabled(
            raw.get("long_pile_delivery_enabled")
        )
        long_pile_delivery = coerce_long_pile_delivery(raw.get("long_pile_delivery_json"))
        append_batches = raw.get("append_batches")

        if kind == "pdf":
            buffer = await asyncio.to_thread(
                generate_commercial_offer_pdf,
                order_data,
                offer_number,
                offer_date,
                customer_name=customer_name,
                manager_name=manager_name,
                manager_phone=None,
                manager_email=None,
                discount_percent=discount_percent,
                kp_db_id=kp_id,
                logistics_cost=logistics_cost,
                pile_logistics_cost=pile_logistics_cost,
                pile_trip_overrides=pile_trip_overrides,
                delivery_conditions=raw.get("delivery_conditions"),
                payment_conditions=raw.get("payment_conditions"),
                append_batches=append_batches,
                pile_catalog_db_path=self.repository.db_path,
                fbs_lm_delivery_enabled=fbs_lm_delivery_enabled,
                weight_catalog_db_path=self.repository.db_path,
                long_pile_delivery_enabled=long_pile_delivery_enabled,
                long_pile_delivery=long_pile_delivery,
            )
            filename = f"КП_{kp_id}.pdf"
        elif kind in {"xlsx", "xlsx_delivery_in_unit"}:
            buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data,
                offer_number,
                offer_date,
                customer_name=customer_name,
                manager_name=manager_name,
                manager_phone=None,
                manager_email=None,
                discount_percent=discount_percent,
                delivery_conditions=raw.get("delivery_conditions"),
                payment_conditions=raw.get("payment_conditions"),
                kp_db_id=kp_id,
                logistics_cost=logistics_cost,
                pile_logistics_cost=pile_logistics_cost,
                pile_trip_overrides=pile_trip_overrides,
                append_batches=append_batches,
                pile_catalog_db_path=self.repository.db_path,
                embed_delivery_in_unit_price=(kind == "xlsx_delivery_in_unit"),
                fbs_lm_delivery_enabled=fbs_lm_delivery_enabled,
                weight_catalog_db_path=self.repository.db_path,
                long_pile_delivery_enabled=long_pile_delivery_enabled,
                long_pile_delivery=long_pile_delivery,
            )
            filename = (
                f"КП_{kp_id}_с_доставкой_в_цене.xlsx"
                if kind == "xlsx_delivery_in_unit"
                else f"КП_{kp_id}.xlsx"
            )
        elif kind == "schema":
            if plate_order_ctx is None:
                raise ArchiveValidationError(
                    "Plate order context is required for schema generation"
                )
            orders_2d = self._orders_2d_from_kp_info(raw)
            if not orders_2d:
                raise ArchiveValidationError("В КП нет позиций для формирования схемы")
            plate_order = AppPlateOrder.from_orders_2d(orders_2d)
            context = await asyncio.to_thread(
                self.optimization_service.optimize,
                plate_order,
                orders_2d=orders_2d,
                plate_order_ctx=plate_order_ctx,
            )
            if not context.optimization_success:
                raise ArchiveValidationError(
                    context.optimization_error_message or "Не удалось выполнить оптимизацию для схемы"
                )

            filename = f"КП_{kp_id}_schema.pdf"
            target_path = self.outputs_dir / filename

            plate_order_ctx.load_optimization_snapshot(
                optimization_result=context.optimization_result,
                plan_by_load=context.plan_by_load,
                load_to_reinforcement_map=context.load_to_reinforcement_map,
            )
            result = await run_in_order_context(
                plate_order_ctx,
                get_visualize_plan(),
                str(self.outputs_dir),
                plate_order_ctx=plate_order_ctx,
            )

            if not isinstance(result, tuple) or len(result) < 2:
                raise ArchiveValidationError("Не удалось создать схему раскладки")
            schema_path = Path(str(result[1]))
            if not schema_path.exists():
                raise ArchiveValidationError("PDF со схемой не был создан")
            if schema_path.resolve() != target_path.resolve():
                await asyncio.to_thread(shutil.copy2, schema_path, target_path)
            return target_path
        else:
            raise ArchiveValidationError(f"Неподдерживаемый тип файла: {kind}")

        target_path = self.outputs_dir / filename
        await asyncio.to_thread(_write_bytes, target_path, buffer.getvalue())
        return target_path

    # ---------- «Актуальный план» (Gantt) ----------

    async def build_current_plan_gantt(self) -> Path:
        """
        Собирает сводную диаграмму Ганта по всем сохранённым планам.
        Импорт plan_manager отложен, чтобы бэкенд запускался без bot-окружения.
        """
        gantt_data = await asyncio.to_thread(
            PlanDistributionService().get_all_plans_gantt_data,
            PlanRepository(),
        )
        if not gantt_data:
            raise ArchiveValidationError(
                "Нет сохранённых планов для создания диаграммы."
            )

        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=gantt_data["all_tracks"],
            tracks_count=3,
            plate_lookup_exact=gantt_data["plate_lookup_exact"],
            plate_lookup_by_length=gantt_data["plate_lookup_by_length"],
            output_dir=str(self.outputs_dir),
            start_date=gantt_data["earliest_start_date"],
        )
        if not gantt_path or not os.path.exists(gantt_path):
            raise ArchiveError("Не удалось создать диаграмму Ганта")
        return Path(gantt_path)

    # ---------- helpers ----------

    @staticmethod
    def _resolve_product_types(raw: dict) -> list[str]:
        """Concrete types for archive list badges (Q3 / MNA-602)."""
        explicit = raw.get("product_types")
        if isinstance(explicit, list) and explicit:
            return [str(t) for t in explicit if str(t).strip() and str(t) != "mixed"]

        product_type = str(raw.get("product_type") or "plates").strip().lower() or "plates"
        if product_type != "mixed":
            return [product_type]

        # Mixed mock/detail payloads may carry typed line arrays without product_types.
        derived: list[str] = []
        for key, _table in (
            ("plates", "kp_plates"),
            ("piles", "kp_piles"),
            ("steps", "kp_steps"),
            ("marches", "kp_marches"),
            ("bridge_piles", "kp_bridge_piles"),
            ("fbs", "kp_fbs"),
        ):
            if raw.get(key):
                derived.append(key)
        return derived

    @staticmethod
    def _optional_counterparty_id(raw: object) -> int | None:
        if raw is None or raw == "":
            return None
        try:
            parsed = int(raw)
        except (TypeError, ValueError):
            return None
        return parsed if parsed >= 1 else None

    def _to_list_item(self, raw: dict) -> ArchiveOfferListItem:
        kp_id = int(raw.get("kp_id") or 0)
        status = raw.get("status") or None
        completion = None
        sgp_progress = None
        shipped_progress = None
        if status in ("в работе", "выполнено", "На СГП"):
            try:
                completion = float(
                    self.repository.get_completion_percentage(kp_id).get("percentage", 0.0)
                )
            except Exception:
                logger.exception("Ошибка получения %% выполнения для КП %s", kp_id)
                completion = None
            try:
                from app.services.sgp_service import SgpService

                progress = SgpService(db_path=self.repository.db_path).sgp_progress(kp_id)
                sgp_progress = {"n": progress.n, "m": progress.m}
            except Exception:
                logger.exception("Ошибка получения sgp_progress для КП %s", kp_id)
            try:
                shipped_progress = self._shipped_progress(kp_id)
            except Exception:
                logger.exception("Ошибка получения shipped_progress для КП %s", kp_id)
                shipped_progress = None
        return ArchiveOfferListItem(
            kp_id=kp_id,
            creation_date=raw.get("creation_date"),
            customer_name=raw.get("customer_name"),
            manager_name=raw.get("manager_name"),
            discount_percent=float(raw.get("discount_percent") or 0),
            subtotal=float(raw.get("subtotal") or 0),
            vat_amount=float(raw.get("vat_amount") or 0),
            total_amount=float(raw.get("total_amount") or 0),
            execution_terms=raw.get("execution_terms") or None,
            status=status,
            completion_percentage=completion,
            sgp_progress=sgp_progress,
            shipped_progress=shipped_progress,
            product_type=str(raw.get("product_type") or "plates"),
            product_types=self._resolve_product_types(raw),
            counterparty_id=self._optional_counterparty_id(raw.get("counterparty_id")),
        )

    def _shipped_progress(self, kp_id: int) -> dict[str, int] | None:
        """SHIP-301: x = Σ плит в done-рейсах КП, m = kp_meta.ordered_qty (read-only)."""
        from core.kp_db_common import _connect
        from core.kp_db_shipments import shipped_qty_for_kp

        conn = _connect(self.repository.db_path)
        try:
            cur = conn.cursor()
            x = shipped_qty_for_kp(cur, kp_id)
            cur.execute(
                "SELECT ordered_qty FROM kp_meta WHERE kp_id = ?",
                (kp_id,),
            )
            row = cur.fetchone()
            if row is None or row[0] is None:
                return None
            return {"x": x, "m": int(row[0])}
        finally:
            conn.close()

    def _to_details(self, raw: dict) -> ArchiveOfferDetails:
        plates = [self._plate_item(p) for p in (raw.get("plates") or [])]
        piles = [self._pile_item(p) for p in (raw.get("piles") or [])]
        steps = [self._step_item(s) for s in (raw.get("steps") or [])]
        marches = [self._march_item(m) for m in (raw.get("marches") or [])]
        bridge_piles = [self._bridge_pile_item(b) for b in (raw.get("bridge_piles") or [])]
        fbs = [self._fbs_item(b) for b in (raw.get("fbs") or [])]
        product_type = str(raw.get("product_type") or "plates")
        kp_id = int(raw.get("kp_id") or 0)
        completion = None
        if raw.get("status") in ("в работе", "выполнено", "На СГП"):
            try:
                completion = float(
                    self.repository.get_completion_percentage(kp_id).get("percentage", 0.0)
                )
            except Exception:
                logger.exception("Ошибка получения %% выполнения для КП %s", kp_id)

        order_data = order_data_from_kp_info(raw)
        logistics_cost = max(0.0, float(raw.get("logistics_cost") or 0.0))
        pile_logistics_cost = max(0.0, float(raw.get("pile_logistics_cost") or 0.0))
        from core.commercial_pricing import (
            calculate_total_cost,
            coerce_fbs_lm_delivery_enabled,
            coerce_long_pile_delivery_enabled,
        )
        from core.pile_trip_pricing import coerce_pile_trip_overrides

        fbs_lm_enabled = coerce_fbs_lm_delivery_enabled(raw.get("fbs_lm_delivery_enabled"))
        long_pile_enabled = coerce_long_pile_delivery_enabled(
            raw.get("long_pile_delivery_enabled")
        )
        totals = calculate_total_cost(
            order_data,
            float(raw.get("discount_percent") or 0.0),
            logistics_cost=logistics_cost,
            db_path=self.repository.db_path,
            require_all_priced=False,
            pile_logistics_cost=pile_logistics_cost,
            pile_trip_overrides=coerce_pile_trip_overrides(
                raw.get("pile_trip_overrides_json")
            ),
            pile_catalog_db_path=self.repository.db_path,
            fbs_lm_delivery_enabled=fbs_lm_enabled,
            weight_catalog_db_path=self.repository.db_path,
            long_pile_delivery_enabled=long_pile_enabled,
            long_pile_delivery=raw.get("long_pile_delivery_json"),
            kp_id=kp_id,
        )
        # Delivery / cargo for archive details: plates 18600 + piles hybrid + ФБС/ЛС/ЛМ.
        total_cargo_weight_kg = float(
            total_order_cargo_weight_kg(order_data, product_types={"plates"})
        )
        plate_delivery_total = float(totals.get("plate_delivery_total") or 0.0)
        pile_delivery_total = float(totals.get("pile_delivery_total") or 0.0)
        fbs_lm_delivery_total = float(totals.get("fbs_lm_delivery_total") or 0.0)
        long_pile_delivery_total = float(totals.get("long_pile_delivery_total") or 0.0)
        delivery_total = (
            plate_delivery_total
            + pile_delivery_total
            + fbs_lm_delivery_total
            + long_pile_delivery_total
        )

        readiness = None
        status = raw.get("status") or ""
        if status in ("в работе", "На СГП"):
            try:
                from app.services.kp_readiness_service import KpReadinessService

                readiness = KpReadinessService(db_path=self.repository.db_path).build_summary(
                    kp_id, status=status
                )
            except Exception:
                logger.exception("Ошибка получения readiness для КП %s", kp_id)

        saved_spec, stale_spec = self._specification_flags(raw)
        return ArchiveOfferDetails(
            kp_id=kp_id,
            creation_date=raw.get("creation_date"),
            customer_name=raw.get("customer_name"),
            customer_inn=raw.get("customer_inn"),
            customer_kpp=raw.get("customer_kpp"),
            counterparty_id=self._optional_counterparty_id(raw.get("counterparty_id")),
            manager_name=raw.get("manager_name"),
            status=raw.get("status"),
            execution_terms=raw.get("execution_terms") or None,
            delivery_conditions=raw.get("delivery_conditions") or None,
            payment_conditions=raw.get("payment_conditions") or None,
            finance=ArchiveOfferFinance(
                subtotal=float(raw.get("subtotal") or 0),
                vat_amount=float(raw.get("vat_amount") or 0),
                total_amount=float(raw.get("total_amount") or 0),
                discount_percent=float(raw.get("discount_percent") or 0),
            ),
            logistics_cost=logistics_cost,
            pile_logistics_cost=pile_logistics_cost,
            pile_trip_overrides=coerce_pile_trip_overrides(
                raw.get("pile_trip_overrides_json")
            ),
            pile_trips=int(totals.get("pile_trips") or 0),
            pile_trip_pending_marks=list(totals.get("pile_trip_pending_marks") or []),
            pile_delivery_ready=bool(totals.get("pile_delivery_ready", True)),
            plate_delivery_total=plate_delivery_total,
            pile_delivery_total=pile_delivery_total,
            fbs_lm_delivery_total=fbs_lm_delivery_total,
            fbs_lm_cargo_kg=float(totals.get("fbs_lm_cargo_kg") or 0.0),
            fbs_lm_trips=int(totals.get("fbs_lm_trips") or 0),
            fbs_lm_delivery_ready=bool(totals.get("fbs_lm_delivery_ready", True)),
            fbs_lm_pending_marks=list(totals.get("fbs_lm_pending_marks") or []),
            fbs_lm_delivery_enabled=fbs_lm_enabled,
            long_pile_delivery_enabled=long_pile_enabled,
            long_pile_delivery_total=long_pile_delivery_total,
            long_pile_lengths=list(totals.get("long_pile_lengths") or []),
            long_pile_pending_marks=list(totals.get("long_pile_pending_marks") or []),
            total_cargo_weight_kg=total_cargo_weight_kg,
            delivery_service_total_rub=delivery_total,
            product_type=product_type,
            plates=plates,
            piles=piles,
            steps=steps,
            marches=marches,
            bridge_piles=bridge_piles,
            fbs=fbs,
            completion_percentage=completion,
            readiness=readiness,
            order_number_1c=(raw.get("order_number_1c") or None),
            paid_at=raw.get("paid_at") or None,
            correction_pending=self._correction_pending(raw),
            invoice_export_block=self._invoice_export_block(raw),
            invoice_warehouse=_stored_invoice_warehouse(raw.get("invoice_warehouse")),
            specification_saved=saved_spec,
            specification_stale_custom=stale_spec,
        )

    @staticmethod
    def _march_item(raw: dict) -> ArchiveMarchItem:
        return ArchiveMarchItem(
            position_number=raw.get("position_number"),
            mark=raw.get("mark") or "",
            concrete_grade=raw.get("concrete_grade") or "",
            qty=int(raw.get("qty") or 0),
            unit_price=_nullable_float(raw.get("unit_price")),
            discounted_price=_nullable_float(raw.get("discounted_price")),
            **_concrete_spec_kwargs(raw),
        )

    @staticmethod
    def _bridge_pile_item(raw: dict) -> ArchiveBridgePileItem:
        return ArchiveBridgePileItem(
            position_number=raw.get("position_number"),
            mark=raw.get("mark") or "",
            concrete_grade=raw.get("concrete_grade") or "",
            qty=int(raw.get("qty") or 0),
            unit_price=_nullable_float(raw.get("unit_price")),
            discounted_price=_nullable_float(raw.get("discounted_price")),
            **_concrete_spec_kwargs(raw),
        )

    @staticmethod
    def _fbs_item(raw: dict) -> ArchiveFbsItem:
        return ArchiveFbsItem(
            position_number=raw.get("position_number"),
            mark=raw.get("mark") or "",
            concrete_grade=raw.get("concrete_grade") or "",
            qty=int(raw.get("qty") or 0),
            unit_price=_nullable_float(raw.get("unit_price")),
            discounted_price=_nullable_float(raw.get("discounted_price")),
            **_concrete_spec_kwargs(raw),
        )

    @staticmethod
    def _pile_item(raw: dict) -> ArchivePileItem:
        return ArchivePileItem(
            position_number=raw.get("position_number"),
            mark=raw.get("mark") or "",
            concrete_grade=raw.get("concrete_grade") or "",
            qty=int(raw.get("qty") or 0),
            unit_price=_nullable_float(raw.get("unit_price")),
            discounted_price=_nullable_float(raw.get("discounted_price")),
            **_concrete_spec_kwargs(raw),
        )

    @staticmethod
    def _step_item(raw: dict) -> ArchiveStepItem:
        return ArchiveStepItem(
            position_number=raw.get("position_number"),
            mark=raw.get("mark") or "",
            qty=int(raw.get("qty") or 0),
            unit_price=_nullable_float(raw.get("unit_price")),
            discounted_price=_nullable_float(raw.get("discounted_price")),
        )

    @staticmethod
    def _plate_item(raw: dict) -> ArchivePlateItem:
        return ArchivePlateItem(
            id=_nullable_int(raw.get("id")),
            position_number=raw.get("position_number"),
            plate_name=raw.get("plate_name") or "",
            length_m=_nullable_float(raw.get("length_m")),
            width_m=_nullable_float(raw.get("width_m")),
            load_class=_nullable_int(raw.get("load_class")),
            qty=int(raw.get("qty") or 0),
            unit_price=_nullable_float(raw.get("unit_price")),
            discounted_price=_nullable_float(raw.get("discounted_price")),
            unit_weight=_nullable_float(raw.get("unit_weight")),
            total_weight=_nullable_float(raw.get("total_weight")),
            status=raw.get("status") or None,
            **_concrete_spec_kwargs(raw),
        )

    @staticmethod
    def _orders_2d_from_kp_info(kp_info: dict) -> list[dict]:
        plates = kp_info.get("plates") or []
        result: list[dict] = []
        for plate in plates:
            length_m = float(plate.get("length_m") or 0)
            width_m = float(plate.get("width_m") or 1.2)
            qty = int(plate.get("qty") or 0)
            if length_m <= 0 or qty <= 0:
                continue
            load_class = int(plate.get("load_class") or 800)
            load_code = max(1, load_class // 100)
            result.append(
                {
                    "length": length_m,
                    "width": int(round(width_m * 1000)),
                    "qty": qty,
                    "load_code": load_code,
                    "length_dm_raw": "",
                    "plate_name": plate.get("plate_name") or "",
                    "kp_id": kp_info.get("kp_id"),
                }
            )
        return result

    def _resolve_target_iso(self, raw: dict, target: str | None) -> str:
        if target:
            try:
                return date.fromisoformat(target).isoformat()
            except ValueError as exc:
                raise ArchiveValidationError(
                    "Параметр target должен быть ISO-датой YYYY-MM-DD"
                ) from exc
        terms = (raw.get("execution_terms") or "").strip()
        if not terms:
            raise ArchiveValidationError(
                "Укажите target или сохраните срок изготовления в КП"
            )
        formatted = self._parse_execution_terms(terms)
        return datetime.strptime(formatted, "%d.%m.%Y").date().isoformat()

    def _build_capacity_snapshot(self, raw: dict, *, target_iso: str) -> CapacitySnapshot:
        today = self._resolve_today()
        items = self._plates_to_batch_items(raw)
        occupancy = self._load_occupancy()
        workdays = self._collect_workdays(today_iso=today, target_iso=target_iso)
        holidays = sorted(d.isoformat() for d in load_holidays())
        extra = sorted(d.isoformat() for d in load_extra_workdays())
        return build_capacity_snapshot(
            items=items,
            target_date=target_iso,
            occupancy=occupancy,
            workdays=workdays,
            produced={},
            today=today,
            holidays=holidays,
            extra_workdays=extra,
        )

    def _resolve_today(self) -> str:
        if self._today_override is not None:
            return date.fromisoformat(self._today_override).isoformat()
        return date.today().isoformat()

    @staticmethod
    def _plates_to_batch_items(raw: dict) -> list[BatchItemInput]:
        items: list[BatchItemInput] = []
        for idx, plate in enumerate(raw.get("plates") or []):
            plate_id = int(plate.get("id") or plate.get("plate_id") or idx + 1)
            qty = int(plate.get("qty") or 0)
            length_m = float(plate.get("length_m") or 0)
            if qty <= 0 or length_m <= 0:
                continue
            items.append(
                BatchItemInput(plate_id=plate_id, qty=qty, length_m=length_m)
            )
        return items

    @staticmethod
    def _load_occupancy() -> dict[str, dict]:
        try:
            calendar = PlanDistributionService().get_global_calendar_info(
                PlanRepository()
            )
        except Exception:
            logger.exception("capacity gate: occupancy unavailable")
            return {}
        if not calendar:
            return {}
        days_info = calendar.get("days_info") or {}
        return days_info if isinstance(days_info, dict) else {}

    @staticmethod
    def _collect_workdays(*, today_iso: str, target_iso: str) -> set[str]:
        today_d = date.fromisoformat(today_iso)
        end = max(
            today_d + timedelta(days=400),
            date.fromisoformat(target_iso) + timedelta(days=60),
        )
        holidays = load_holidays()
        extra_workdays = load_extra_workdays()
        workdays: set[str] = set()
        current = today_d
        while current <= end:
            if is_working_day(current, holidays, extra_workdays):
                workdays.add(current.isoformat())
            current += timedelta(days=1)
        return workdays

    @staticmethod
    def _snapshot_to_response(snap: CapacitySnapshot) -> CapacitySnapshotResponse:
        days: dict[str, CapacityDayInfo] = {}
        for key, info in (snap.days_info or {}).items():
            if not isinstance(info, dict):
                continue
            days[key] = CapacityDayInfo(
                occupied=float(info.get("occupied", 0) or 0),
                max=float(info.get("max", TRACKS_PER_DAY_DEFAULT) or TRACKS_PER_DAY_DEFAULT),
            )
        return CapacitySnapshotResponse(
            start_date=snap.start_date,
            target_date=snap.target_date,
            tracks_needed=snap.tracks_needed,
            tracks_free_in_window=snap.tracks_free_in_window,
            delta=snap.delta,
            status=snap.status,
            hint=snap.hint,
            days_info=days,
            holidays=list(snap.holidays),
            extra_workdays=list(snap.extra_workdays),
            calendar_from_month=snap.calendar_from_month,
            calendar_to_month=snap.calendar_to_month,
        )

    @staticmethod
    def _parse_execution_terms(raw: str) -> str:
        try:
            formatted, _ = parse_execution_terms(raw, policy="strict")
            return formatted
        except ValueError as exc:
            raise ArchiveValidationError(str(exc)) from exc


def _specification_filename(kp_id: int, invoice_number: str | None) -> str:
    number = (invoice_number or "").strip()
    if not number:
        return f"Спецификация КП {kp_id}.xlsx"
    safe = number.replace("/", "-").replace("\\", "-").replace("\x00", "")
    return f"Спецификация по счету {safe}.xlsx"


def _short_person_name(full_name: str) -> str:
    parts = [part for part in full_name.split() if part]
    if len(parts) < 2:
        return full_name.strip()
    initials = " ".join(f"{part[0]}." for part in parts[1:])
    return f"{parts[0]} {initials}".strip()


def _parse_contract_date(value: object) -> date:
    text = str(value or "").strip()
    if not text:
        raise ArchiveValidationError("В договоре нет даты")
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        pass
    try:
        return datetime.strptime(text, "%d.%m.%Y").date()
    except ValueError as exc:
        raise ArchiveValidationError("В договоре некорректная дата") from exc


def _load_stored_specification(raw: dict) -> dict | None:
    text = raw.get("specification_json")
    if text is None or not str(text).strip():
        return None
    try:
        document = json.loads(text)
    except json.JSONDecodeError:
        return None
    if not isinstance(document, dict):
        return None
    return document


def _choice_from_payload(payload: SpecificationChoiceIn) -> SpecificationChoice:
    try:
        return SpecificationChoice(
            payment=payload.payment,
            term=payload.term,
            delivery=payload.delivery,
            payment_date=_payload_date(payload.payment_date),
            payment_days=payload.payment_days,
            second_share_percent=payload.second_share_percent,
            second_payment_date=_payload_date(payload.second_payment_date),
            custom_text=payload.custom_text,
            term_date=_payload_date(payload.term_date),
            pile_count=payload.pile_count,
            pile_unit=payload.pile_unit,
            term_days=payload.term_days,
            delivery_address=payload.delivery_address,
        )
    except SpecificationTextError as exc:
        raise ArchiveValidationError(str(exc)) from exc


def _payload_date(value: str | None) -> date | None:
    if value is None or not str(value).strip():
        return None
    try:
        return date.fromisoformat(str(value).strip())
    except ValueError as exc:
        raise ArchiveValidationError("Некорректная дата") from exc


def _choice_out(document: dict) -> SpecificationChoiceOut:
    return SpecificationChoiceOut(
        payment=document.get("payment"),
        term=document.get("term"),
        delivery=document.get("delivery"),
        payment_date=document.get("payment_date"),
        payment_days=document.get("payment_days"),
        second_share_percent=document.get("second_share_percent"),
        second_payment_date=document.get("second_payment_date"),
        custom_text=document.get("custom_text"),
        term_date=document.get("term_date"),
        pile_count=document.get("pile_count"),
        pile_unit=document.get("pile_unit"),
        term_days=document.get("term_days"),
        delivery_address=document.get("delivery_address"),
    )


def _suggested_choice(raw: dict, delivery_kopecks: int) -> SpecificationChoiceOut:
    payment = "prepay_100" if not str(raw.get("payment_conditions") or "").strip() else None
    delivery = None
    text = str(raw.get("delivery_conditions") or "")
    if "самовывоз" in text.casefold() and delivery_kopecks == 0:
        delivery = "pickup"
    return SpecificationChoiceOut(payment=payment, delivery=delivery)


def _suggestion_paragraphs(choice: SpecificationChoiceOut, total: int) -> tuple[str, str, str]:
    if choice.payment != "prepay_100" and choice.delivery != "pickup":
        return "", "", ""
    rendered = render_specification(
        SpecificationChoice(
            payment="prepay_100",
            term="by_date",
            term_date=date(2000, 1, 1),
            delivery="pickup",
        ),
        payable_total=total,
        has_piles=False,
    )
    payment = rendered.payment if choice.payment == "prepay_100" else ""
    delivery = rendered.delivery if choice.delivery == "pickup" else ""
    return payment, "", delivery


def _kp_has_piles(raw: dict) -> bool:
    return any(isinstance(item, dict) for item in (raw.get("piles") or []))


def _uniform_concrete_grade(raw: dict) -> str | None:
    piles = [item for item in (raw.get("piles") or []) if isinstance(item, dict)]
    if not piles:
        return None
    grades: list[str] = []
    for pile in piles:
        grade = str(pile.get("concrete_grade") or "").strip()
        if not grade:
            return None
        grades.append(grade)
    if len(set(grades)) != 1:
        return None
    return grades[0]


def _rub_kopecks(amount: object) -> int:
    value = Decimal(str(amount or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return int((value * Decimal(100)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _concrete_spec_kwargs(raw: dict) -> dict[str, str | None]:
    def text(key: str) -> str | None:
        value = raw.get(key)
        if value is None:
            return None
        stripped = str(value).strip()
        return stripped or None

    source = text("concrete_spec_source")
    aggregate_raw = raw.get("concrete_aggregate")
    if aggregate_raw is None:
        aggregate = None
    else:
        aggregate_text = str(aggregate_raw)
        aggregate = "" if source == "manual" and aggregate_text.strip() == "" else aggregate_text.strip() or None
    return {
        "frost_resistance": text("frost_resistance"),
        "waterproofness": text("waterproofness"),
        "concrete_aggregate": aggregate,
        "concrete_spec_source": source,
    }


def _invoice_order_ids(raw: dict) -> tuple[str, str, str] | None:
    number = str(raw.get("order_number_1c") or "").strip()
    uid_order = str(raw.get("uid_order_1c") or "").strip()
    uid_kp = str(raw.get("uid_kp_1c") or "").strip()
    if not number or not uid_order or not uid_kp:
        return None
    return number, uid_order, uid_kp


def _stored_invoice_warehouse(value: object) -> str | None:
    text = str(value or "").strip()
    return text or None


def _require_invoice_warehouse(raw: str | None) -> str:
    name = str(raw or "").strip()
    if not name:
        raise ArchiveValidationError("Укажите склад")
    if name not in INVOICE_WAREHOUSES:
        raise ArchiveValidationError("Неизвестный склад")
    return name


def _warehouse_for_correction(raw: dict, requested: str | None) -> str:
    stored = _stored_invoice_warehouse(raw.get("invoice_warehouse"))
    if stored is None:
        return _require_invoice_warehouse(requested)
    if stored not in INVOICE_WAREHOUSES:
        raise ArchiveValidationError("Неизвестный склад")
    return stored


_INVOICE_LINE_GROUPS = (
    ("plates", "plate", False),
    ("piles", "pile", True),
    ("steps", "stair_step", False),
    ("marches", "stair_flight", False),
    ("bridge_piles", "bridge_pile", True),
    ("fbs", "fbs", False),
)


def _allocation_rows(
    raw: dict,
    *,
    long_enabled: bool,
    catalog_lookup,
) -> tuple[list[int], list[str], list[int | None]]:
    qtys: list[int] = []
    types: list[str] = []
    keys: list[int | None] = []
    for group, kind, is_pile in _INVOICE_LINE_GROUPS:
        for item in raw.get(group) or []:
            if not isinstance(item, dict):
                continue
            qtys.append(_non_negative_qty(item.get("qty")))
            types.append(kind)
            if is_pile and long_enabled and catalog_lookup is not None:
                keys.append(_long_pile_length_key(item, catalog_lookup))
            else:
                keys.append(None)
    return qtys, types, keys


def _non_negative_qty(raw: object) -> int:
    try:
        return max(0, int(raw or 0))
    except (TypeError, ValueError):
        return 0


def _long_pile_length_key(item: dict, catalog_lookup) -> int | None:
    from core.pile_catalog import parse_bridge_pile_geometry, parse_pile_mark
    from core.pile_trip_pricing import LONG_PILE_LENGTH_M_MIN

    mark = str(item.get("mark") or item.get("name") or "").strip()
    entry = catalog_lookup(mark) if mark else None
    length_m = getattr(entry, "length_m", None) if entry is not None else None
    if length_m is None and mark:
        length_m, _section = parse_bridge_pile_geometry(mark)
        if length_m is None:
            length_m, _section = parse_pile_mark(mark)
    if length_m is None or float(length_m) <= LONG_PILE_LENGTH_M_MIN:
        return None
    return int(round(float(length_m) * 10))


def _ready_long_pile_amounts(totals: dict) -> dict[int, float]:
    if not totals.get("long_pile_delivery_enabled"):
        return {}
    amounts: dict[int, float] = {}
    for row in totals.get("long_pile_lengths") or []:
        if not isinstance(row, dict) or not row.get("ready"):
            continue
        amount = float(row.get("amount") or 0.0)
        if amount <= 0:
            continue
        amounts[int(row.get("length_key") or 0)] = amount
    return amounts


def _delivery_kopecks(amount: float) -> int:
    return int(round(float(amount) * 100))


def _expected_delivery_kopecks(totals: dict, long_amounts: dict[int, float]) -> int:
    return (
        _delivery_kopecks(float(totals.get("plate_delivery_total") or 0.0))
        + _delivery_kopecks(float(totals.get("pile_delivery_total") or 0.0))
        + _delivery_kopecks(float(totals.get("fbs_lm_delivery_total") or 0.0))
        + sum(_delivery_kopecks(amount) for amount in long_amounts.values())
    )


def _unready_delivery_reason(totals: dict) -> str | None:
    from core.pile_trip_pricing import format_long_pile_delivery_label

    parts: list[str] = []
    if not bool(totals.get("pile_delivery_ready", True)):
        marks = _mark_list(totals.get("pile_trip_pending_marks"))
        parts.append(f"Котёл доставки свай не готов: нет веса или тарифа у марок {marks}")
    if not bool(totals.get("fbs_lm_delivery_ready", True)):
        marks = _mark_list(totals.get("fbs_lm_pending_marks"))
        parts.append(
            "Котёл доставки ФБС, ступеней и маршей не готов: нет веса или тарифа у марок "
            f"{marks}"
        )
    if totals.get("long_pile_delivery_enabled"):
        for row in totals.get("long_pile_lengths") or []:
            if not isinstance(row, dict) or row.get("ready"):
                continue
            label = format_long_pile_delivery_label(int(row.get("length_key") or 0))
            marks = _mark_list(row.get("pending_marks"))
            if marks:
                parts.append(f"{label} не готова: нет веса или тарифа у марок {marks}")
            else:
                parts.append(f"{label} не готова: нет тарифа")
    if not parts:
        return None
    return "Доставка не посчитана. " + "; ".join(parts)


def _mark_list(raw: object) -> str:
    if not isinstance(raw, (list, tuple)):
        return ""
    return ", ".join(str(mark) for mark in raw if str(mark).strip())


def _nullable_float(value: object) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _nullable_int(value: object) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
