from __future__ import annotations

import asyncio
import logging
import os
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Literal

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
)
from app.repositories.plan_repository import PlanRepository
from app.services.capacity_gate_service import (
    CapacityGateBlockedError,
    CapacitySnapshot,
    assert_capacity_allows_save,
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


class ArchiveService:
    """Бизнес-логика раздела «Архив» (аналог bot/handlers/archive.py)."""

    def __init__(
        self,
        repository: KpArchiveRepository | None = None,
        outputs_dir: Path | None = None,
        optimization_service: OptimizationService | None = None,
        file_generation_service: FileGenerationService | None = None,
    ) -> None:
        settings = get_settings()
        self.repository = repository or KpArchiveRepository()
        self.outputs_dir = outputs_dir or settings.outputs_dir
        self.outputs_dir.mkdir(parents=True, exist_ok=True)
        self.optimization_service = optimization_service or OptimizationService()
        self.file_generation_service = file_generation_service or FileGenerationService()
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
        """Hydrate a commercial draft from archive KP (status «в архиве» only)."""
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
    ) -> ArchiveOfferDetails:
        """Обновляет тарифы рейсов и суммы заказа (без авто-PDF)."""
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        trip = max(0.0, float(logistics_cost or 0.0))
        if not self.repository.update_logistics_cost(
            kp_id,
            trip,
            pile_logistics_cost=pile_logistics_cost,
            pile_trip_overrides=pile_trip_overrides,
        ):
            raise ArchiveNotFoundError(
                f"Не удалось обновить стоимость рейса. КП №{kp_id} не найдено или пустое."
            )
        return self.get_details(kp_id, user=user)

    def delete_offer(self, kp_id: int, *, user: dict) -> None:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} уже удалено или не существует")
        assert_offer_write_access(user, raw)
        if not self.repository.delete(kp_id):
            raise ArchiveNotFoundError(f"КП №{kp_id} уже удалено или не существует")

    def move_to_production(self, kp_id: int, terms_input: str, *, user: dict) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        if raw.get("status") != "в архиве":
            raise ArchiveValidationError(
                "Перевести в производство можно только КП из раздела «в архиве»"
            )

        execution_terms = self._parse_execution_terms(terms_input)
        self._enforce_capacity_gate_for_terms(raw, execution_terms)
        try:
            from core.kp.offers_write import commit_move_to_production

            commit_move_to_production(kp_id, execution_terms, self.repository.db_path)
        except Exception as exc:
            logger.exception("move_to_production failed for kp_id=%s", kp_id)
            raise ArchiveError(
                f"Не удалось перевести КП №{kp_id} в производство"
            ) from exc

        return self.get_details(kp_id, user=user)

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
            )
            filename = f"КП_{kp_id}.pdf"
        elif kind == "xlsx":
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
            )
            filename = f"КП_{kp_id}.xlsx"
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
        from core.commercial_pricing import calculate_total_cost
        from core.pile_trip_pricing import coerce_pile_trip_overrides

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
        )
        # Delivery / cargo for archive details: plates 18600 + piles hybrid.
        total_cargo_weight_kg = float(
            total_order_cargo_weight_kg(order_data, product_types={"plates"})
        )
        plate_delivery_total = float(totals.get("plate_delivery_total") or 0.0)
        pile_delivery_total = float(totals.get("pile_delivery_total") or 0.0)
        delivery_total = plate_delivery_total + pile_delivery_total

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

        return ArchiveOfferDetails(
            kp_id=kp_id,
            creation_date=raw.get("creation_date"),
            customer_name=raw.get("customer_name"),
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

    def _enforce_capacity_gate_for_terms(self, raw: dict, execution_terms_ddmmyyyy: str) -> None:
        target_iso = datetime.strptime(execution_terms_ddmmyyyy, "%d.%m.%Y").date().isoformat()
        snap = self._build_capacity_snapshot(raw, target_iso=target_iso)
        try:
            assert_capacity_allows_save(snap)
        except CapacityGateBlockedError as exc:
            raise ArchiveValidationError(str(exc)) from exc

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
