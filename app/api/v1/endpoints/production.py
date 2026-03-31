from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException

from app.dependencies.auth import require_roles
from app.schemas.production import CompleteProductionDayRequest, CreatePlanRequest, SaveWorkCalendarRequest
from app.services.production_service import ProductionService

router = APIRouter(prefix="/production", tags=["production"])


@router.get("/plans")
def list_plans(
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    return ProductionService().list_plans()


@router.post("/plans")
def create_plan(
    payload: CreatePlanRequest,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    return ProductionService().create_plan(**payload.model_dump())


@router.get("/plans/{plan_id}")
def get_plan(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    plan = ProductionService().get_plan(plan_id)
    if not plan:
        raise HTTPException(status_code=404, detail="Plan not found")
    return plan


@router.post("/plans/{plan_id}/activate")
def activate_plan(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    return ProductionService().activate_plan(plan_id)


@router.get("/calendar")
def get_calendar(
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    calendar = ProductionService().get_calendar()
    if not calendar:
        return {"plans_count": 0, "days_info": {}, "completed_days": []}
    return calendar


@router.get("/days/{target_date}")
def get_day_view(
    target_date: str,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    data = ProductionService().get_day_view(target_date)
    if not data:
        raise HTTPException(status_code=404, detail="Day not found")
    return data


@router.post("/days/{target_date}/complete")
def complete_day(
    target_date: str,
    payload: CompleteProductionDayRequest,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    return ProductionService().complete_day(plan_id=payload.plan_id, target_date=target_date)


@router.get("/candidates")
def list_candidates(
    limit: int = 500,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    items = ProductionService().load_candidates_for_plan(limit=limit)
    return {"items": items, "count": len(items)}


@router.get("/work-calendar")
def get_work_calendar(
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    return ProductionService().get_work_calendar()


@router.put("/work-calendar")
def save_work_calendar(
    payload: SaveWorkCalendarRequest,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    return ProductionService().save_work_calendar(payload.model_dump())

