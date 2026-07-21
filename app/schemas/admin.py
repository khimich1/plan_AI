from __future__ import annotations

from pydantic import BaseModel, Field


class DbStatsResponse(BaseModel):
    """Сводная статистика по состоянию БД и хранилищу планов."""

    kp_total: int = Field(..., description="Всего сохранённых КП в plita.db")
    kp_in_work: int = Field(..., description="КП со статусом 'в работе'")
    kp_completed: int = Field(..., description="КП со статусом 'выполнено'")
    plates_in_work: int = Field(..., description="Записей плит в kp_plates")
    plates_completed: int = Field(..., description="Записей в completed_plates")
    plate_rests: int = Field(..., description="Записей в plate_rests")
    plans_count: int = Field(..., description="Количество JSON-планов в data/plans")
    current_plan_present: bool = Field(
        ..., description="Существует ли файл current_plan.json"
    )


class DbResetReport(BaseModel):
    """Унифицированный отчёт по выполненной операции обнуления."""

    sqlite: dict[str, int] = Field(
        default_factory=dict,
        description="Счётчики удалённых записей по таблицам plita.db",
    )
    plans: dict[str, int] = Field(
        default_factory=dict,
        description="Счётчики удалённых файлов планов (current_plan, metadata, plan_files, total)",
    )
    calendar_reset: bool = Field(
        default=False,
        description="True, если файл work_calendar.json был приведён к пустому состоянию",
    )


class RecoverPlatesResponse(BaseModel):
    """Результат восстановления 'застрявших' плит."""

    recovered_records: int = Field(
        ..., description="Сколько записей kp_plates переведено из 'в плане' в 'в производстве'"
    )
