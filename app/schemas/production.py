from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any, Literal
from zoneinfo import ZoneInfo

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from core.production.capacity import TRACKS_PER_DAY_HARD_CAP

# D5: inclusive span ≤ 366 calendar days ⇒ (max − min).days ≤ 365.
# Forward horizon: min_fill ≤ today+366. No floor into the past.
_FILL_SPAN_MAX_DELTA_DAYS = 365
_FILL_FORWARD_HORIZON_DAYS = 366
_MOSCOW_TZ = ZoneInfo("Europe/Moscow")


def _validate_fill_targets_horizon(
    items: list[FillTargetItem],
) -> list[FillTargetItem]:
    """Reject fill lists whose date span or forward horizon exceeds D5 limits."""
    if not items:
        return items
    days = [item.date for item in items]
    min_d = min(days)
    max_d = max(days)
    if (max_d - min_d).days > _FILL_SPAN_MAX_DELTA_DAYS:
        raise ValueError(
            "Диапазон дат fill_targets не может превышать 366 дней "
            f"(сейчас {(max_d - min_d).days + 1})."
        )
    today = datetime.now(_MOSCOW_TZ).date()
    if min_d > today + timedelta(days=_FILL_FORWARD_HORIZON_DAYS):
        raise ValueError(
            "Минимальная дата fill_targets не может быть дальше "
            f"сегодня+{_FILL_FORWARD_HORIZON_DAYS} дней."
        )
    return items


class CreatePlanRequest(BaseModel):
    name: str = Field(default="Новый план")
    start_date: str
    tracks_per_day: int = Field(ge=1, le=50)
    all_tracks_list: list = Field(default_factory=list)
    plate_lookup_exact: dict = Field(default_factory=dict)
    plate_lookup_by_length: dict = Field(default_factory=dict)
    orders_2d: list = Field(default_factory=list)
    optimization_result: dict = Field(default_factory=dict)
    active_plan_id: str | None = None


class RejectedPlateItem(BaseModel):
    """Позиция плиты, которую не нужно списывать как выполненную."""

    track_number: int = Field(ge=1)
    plate_index: int = Field(ge=0)
    qty: int = Field(ge=0)


class CompleteProductionDayRequest(BaseModel):
    plan_id: str
    rejected_plates: list[RejectedPlateItem] = Field(default_factory=list)
    expected_version: int | None = None


class SaveWorkCalendarRequest(BaseModel):
    extra_holidays: list[str] = Field(default_factory=list)
    extra_workdays: list[str] = Field(default_factory=list)


# ---------- Web-флоу «Производство» ----------


class KpCandidatePlateItem(BaseModel):
    """Плита, доступная для постановки в новый план (статус ``в производстве``)."""

    id: int
    plate_name: str
    length_m: float
    width_m: float
    load_class: int | None = None
    qty: int


class KpCandidateItem(BaseModel):
    kp_id: int
    customer_name: str
    creation_date: str
    execution_terms: str
    total_plates: int
    completed_plates: int
    completion_pct: float
    in_plan_pct: float
    total_length_m: float
    plates: list[KpCandidatePlateItem] = Field(default_factory=list)


class KpCandidatesResponse(BaseModel):
    items: list[KpCandidateItem]
    count: int


class DayOccupancyResponse(BaseModel):
    occupancy: dict[str, int]
    max_per_day: int
    max_by_day: dict[str, int] = Field(default_factory=dict)


class DeletePlanResponse(BaseModel):
    plan_id: str
    deleted: bool


class PlanMetaSummary(BaseModel):
    id: str
    name: str
    created_at: str | None = None
    start_date: str | None = None
    total_days: int | None = None
    tracks_count: int | None = None
    total_tracks: int | None = None
    version: int | None = None


class PlansListResponse(BaseModel):
    plans: list[PlanMetaSummary]
    active_plan_id: str | None = None


class PlanDetailResponse(BaseModel):
    """Полный план с optimistic-lock version для web-клиента."""

    model_config = ConfigDict(extra="allow")

    version: int


class RemoveTrackResponse(BaseModel):
    plan_id: str
    date: str
    track_index: int
    plates_returned: int
    saved_tracks_count: int
    warnings: list[str] | None = None


class FillTargetItem(BaseModel):
    """Один пункт корзины дозаполнения: дата + сколько дорожек туда положить."""

    date: date
    tracks: int = Field(ge=1, le=TRACKS_PER_DAY_HARD_CAP)


LayoutReinforcementOrder = Literal["asc", "desc"]


class SgpReservationItem(BaseModel):
    """Закрытие потребности со склада СГП (не уходит в оптимизатор)."""

    sgp_id: int = Field(ge=1)
    target_kp_id: int = Field(ge=1)
    qty: int = Field(ge=1)


class BuildPlanRequest(BaseModel):
    start_date: str
    tracks_count: int = Field(ge=1, le=50)
    filter_method: Literal["all", "kp"]
    selected_kp_ids: list[int] = Field(default_factory=list)
    selected_plate_ids: dict[int, list[int]] = Field(default_factory=dict)
    selected_plate_qty: dict[int, dict[int, int]] = Field(default_factory=dict)
    active_plan_id: str | None = None
    plan_name: str | None = None
    fill_targets: list[FillTargetItem] | None = None
    layout_reinforcement_order: LayoutReinforcementOrder = "asc"
    sgp_reservations: list[SgpReservationItem] = Field(default_factory=list)

    @field_validator("selected_kp_ids")
    @classmethod
    def _unique_kp_ids(cls, value: list[int]) -> list[int]:
        return list(dict.fromkeys(value))

    @field_validator("selected_plate_qty")
    @classmethod
    def _validate_selected_plate_qty(
        cls, value: dict[int, dict[int, int]]
    ) -> dict[int, dict[int, int]]:
        for kp_id, plates in value.items():
            for plate_id, qty in plates.items():
                if int(qty) < 1:
                    raise ValueError(
                        f"Количество плиты #{plate_id} (КП #{kp_id}) должно быть не меньше 1"
                    )
        return value

    @field_validator("fill_targets")
    @classmethod
    def _unique_fill_target_dates(
        cls, value: list[FillTargetItem] | None
    ) -> list[FillTargetItem] | None:
        # Дублирующиеся даты ломали бы _build_tracks_by_day_from_targets
        # (одна и та же дата перезаписалась бы), поэтому отсекаем заранее.
        if value is None:
            return value
        seen: set[date] = set()
        for item in value:
            if item.date in seen:
                raise ValueError(f"Дата {item.date.isoformat()} указана в fill_targets дважды")
            seen.add(item.date)
        return value

    @model_validator(mode="after")
    def _fill_targets_horizon(self) -> BuildPlanRequest:
        if self.fill_targets is not None:
            _validate_fill_targets_horizon(self.fill_targets)
        return self


class BuildPlanSummary(BaseModel):
    total_tracks: int
    total_days: int
    selected_plates_count: int
    kp_count: int


class BuildPlanResponse(BaseModel):
    plan: dict[str, Any]
    stats: dict[str, Any]
    summary: BuildPlanSummary


# ---------- Детальный просмотр дня ----------


class DayPlateInfo(BaseModel):
    """Отдельная агрегированная позиция (плита) внутри дорожки."""

    customer: str = "неизвестно"
    plate_name: str = ""
    kp_date: str = "неизвестно"
    kp_id: int | None = None
    length_m: float
    width_mm: int
    qty: int
    reinforcement: float = 0.0
    # В БД load_class может быть, например, 1250 -> код 12.5.
    # Поэтому принимаем числовой код как int/float.
    load_code: float | int | None = None
    write_off_completed: bool = Field(
        default=False,
        description="Позиция показана из снимка после списания (completed_plates / журнал).",
    )


class DayTrackDetail(BaseModel):
    """Одна дорожка на выбранную дату с агрегированными плитами."""

    track_number: int
    plan_track_index: int = Field(
        default=0,
        description="0-based индекс дорожки внутри плана (для DELETE API).",
    )
    length: float | None = None
    max_reinforcement: float = 0.0
    label: str | None = None
    source_plan_id: str | None = None
    source_plan_name: str | None = None
    plates_info: list[DayPlateInfo] = Field(default_factory=list)


class DayPlanBlock(BaseModel):
    """Блок одного плана в разрезе одного дня."""

    plan_id: str
    plan_name: str
    completed: bool = False
    tracks: list[DayTrackDetail] = Field(default_factory=list)
    from_sgp: list[dict[str, Any]] = Field(
        default_factory=list,
        description="Позиции плана, закрытые со СГП (без дорожки).",
    )
    from_sgp_qty: int = 0


class DayViewDetailResponse(BaseModel):
    """Типизированный ответ /days/{date} с полным содержимым дня."""

    date: str
    plans: list[DayPlanBlock] = Field(default_factory=list)
    plans_count: int = 0
    total_tracks: int = 0


# ---------- Анализ срочных / подложек ----------


class AnalyzeSubstratesRequest(BaseModel):
    """Запрос аналитического прогона: срочные позиции + подложки + дефицит ёмкости."""

    model_config = ConfigDict(extra="forbid")

    fill_targets: list[FillTargetItem]
    deadline_until: date

    @field_validator("fill_targets")
    @classmethod
    def _unique_fill_target_dates(
        cls, value: list[FillTargetItem]
    ) -> list[FillTargetItem]:
        seen: set[date] = set()
        for item in value:
            if item.date in seen:
                raise ValueError(f"Дата {item.date.isoformat()} указана в fill_targets дважды")
            seen.add(item.date)
        return value

    @model_validator(mode="after")
    def _fill_targets_horizon(self) -> AnalyzeSubstratesRequest:
        _validate_fill_targets_horizon(self.fill_targets)
        return self


class UrgentPositionItem(BaseModel):
    plate_id: int
    kp_id: int
    plate_name: str
    qty_remaining: int
    deadline: date
    deadline_source: str
    deadline_details: list[dict[str, Any]] = Field(default_factory=list)
    conflict: str | None = None


class SubstrateRecommendationItem(BaseModel):
    plate_id: int
    kp_id: int
    plate_name: str
    qty_recommended: int
    under_plate_id: int
    under_kp_id: int
    under_plate_name: str
    needed_by: date
    storage_days: int
    saving_mm: int
    saving_m: float


class CapacityOptionItem(BaseModel):
    action: Literal["bump_fill", "propose_day"]
    date: str
    add_tracks: int
    free: int


class CapacityDeficitItem(BaseModel):
    tracks_needed: int
    tracks_available: int
    tracks_missing: int
    deficit_until: str
    options: list[CapacityOptionItem] = Field(default_factory=list)


class AnalysisMetaItem(BaseModel):
    orders_count: int
    analysis_duration_ms: int
    optimization_status: Literal["ok", "partial", "error"]
    error_message: str | None = None


class AnalyzeSubstratesResponse(BaseModel):
    urgent_positions: list[UrgentPositionItem] = Field(default_factory=list)
    substrate_recommendations: list[SubstrateRecommendationItem] = Field(
        default_factory=list
    )
    capacity_deficit: CapacityDeficitItem | None = None
    analysis_meta: AnalysisMetaItem


# ---------- Per-day capacity overrides ----------


class DayCapacityMapResponse(BaseModel):
    """Ёмкость (max_tracks) по каждой дате диапазона: override или default."""

    model_config = ConfigDict(extra="forbid")

    capacity: dict[str, int]


class SaveDayCapacityRequest(BaseModel):
    """Сохранение переопределения ёмкости на день."""

    model_config = ConfigDict(extra="forbid")

    date: str
    max_tracks: int = Field(ge=0, le=TRACKS_PER_DAY_HARD_CAP)

    @field_validator("date")
    @classmethod
    def _valid_iso_date(cls, value: str) -> str:
        try:
            return date.fromisoformat(value).isoformat()
        except ValueError as exc:
            raise ValueError("date must be YYYY-MM-DD") from exc


class SaveDayCapacityResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")

    date: str
    max_tracks: int
