from __future__ import annotations

from pydantic import BaseModel, Field


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


class CompleteProductionDayRequest(BaseModel):
    plan_id: str


class SaveWorkCalendarRequest(BaseModel):
    extra_holidays: list[str] = Field(default_factory=list)
    extra_workdays: list[str] = Field(default_factory=list)

