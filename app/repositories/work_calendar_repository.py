from __future__ import annotations

import json
from datetime import date

from app.core.settings import get_settings
from core.work_calendar import is_working_day, load_extra_workdays, load_holidays, nth_working_day


class WorkCalendarRepository:
    def __init__(self) -> None:
        self.settings = get_settings()
        self.path = self.settings.work_calendar_path

    def load_raw(self) -> dict:
        if not self.path.exists():
            return {"extra_holidays": [], "extra_workdays": []}
        with open(self.path, "r", encoding="utf-8") as file:
            data = json.load(file)
        if not isinstance(data, dict):
            return {"extra_holidays": [], "extra_workdays": []}
        return {
            "extra_holidays": list(data.get("extra_holidays", [])),
            "extra_workdays": list(data.get("extra_workdays", [])),
        }

    def save_raw(self, data: dict) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=2)

    def load_holidays(self) -> set[date]:
        return load_holidays()

    def load_extra_workdays(self) -> set[date]:
        return load_extra_workdays()

    def is_working_day(self, day: date) -> bool:
        return is_working_day(day)

    def nth_working_day(self, start_day: date, n: int) -> date:
        return nth_working_day(start_day, n)

