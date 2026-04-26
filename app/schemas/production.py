from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator


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


class SaveWorkCalendarRequest(BaseModel):
    extra_holidays: list[str] = Field(default_factory=list)
    extra_workdays: list[str] = Field(default_factory=list)


# ---------- Web-флоу «Производство» ----------


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


class KpCandidatesResponse(BaseModel):
    items: list[KpCandidateItem]
    count: int


class DayOccupancyResponse(BaseModel):
    occupancy: dict[str, int]
    max_per_day: int


class DeletePlanResponse(BaseModel):
    plan_id: str
    deleted: bool


class BuildPlanRequest(BaseModel):
    start_date: str
    tracks_count: int = Field(ge=1, le=50)
    filter_method: Literal["all", "kp"]
    selected_kp_ids: list[int] = Field(default_factory=list)
    selected_plate_ids: dict[int, list[int]] = Field(default_factory=dict)
    active_plan_id: str | None = None
    plan_name: str | None = None

    @field_validator("selected_kp_ids")
    @classmethod
    def _unique_kp_ids(cls, value: list[int]) -> list[int]:
        return list(dict.fromkeys(value))


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
    load_code: int | None = None


class DayTrackDetail(BaseModel):
    """Одна дорожка на выбранную дату с агрегированными плитами."""

    track_number: int
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


class DayViewDetailResponse(BaseModel):
    """Типизированный ответ /days/{date} с полным содержимым дня."""

    date: str
    plans: list[DayPlanBlock] = Field(default_factory=list)
    plans_count: int = 0
    total_tracks: int = 0
