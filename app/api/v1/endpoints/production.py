from __future__ import annotations

import logging

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, status
from fastapi.responses import FileResponse

from app.dependencies.auth import require_roles
from app.services.day_documents_service import (
    DayDocumentsError,
    generate_day_breakdown,
    generate_day_formovka,
    generate_day_schema,
    make_cleanup_callback,
)
from app.schemas.production import (
    BuildPlanRequest,
    BuildPlanResponse,
    CompleteProductionDayRequest,
    CreatePlanRequest,
    DayOccupancyResponse,
    DayViewDetailResponse,
    DeletePlanResponse,
    KpCandidatesResponse,
    SaveWorkCalendarRequest,
)
from app.services.production_planning_service import ProductionPlanBuildError
from app.services.production_service import ProductionService

logger = logging.getLogger(__name__)

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


@router.post("/plans/build", response_model=BuildPlanResponse)
def build_plan_from_filters(
    payload: BuildPlanRequest,
    _user: dict = Depends(require_roles("admin", "production")),
) -> BuildPlanResponse:
    try:
        result = ProductionService().build_plan_from_filters(
            start_date=payload.start_date,
            tracks_count=payload.tracks_count,
            filter_method=payload.filter_method,
            selected_kp_ids=payload.selected_kp_ids or None,
            selected_plate_ids=payload.selected_plate_ids or None,
            active_plan_id=payload.active_plan_id,
            plan_name=payload.plan_name,
        )
    except ProductionPlanBuildError as exc:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(exc))
    except Exception as exc:
        logger.exception("[production/build] Непредвиденная ошибка: %s", exc)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось построить план производства.",
        ) from exc
    return BuildPlanResponse(**result)


@router.get("/plans/{plan_id}")
def get_plan(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    plan = ProductionService().get_plan(plan_id)
    if not plan:
        raise HTTPException(status_code=404, detail="Plan not found")
    return plan


@router.delete("/plans/{plan_id}", response_model=DeletePlanResponse)
def delete_plan(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
) -> DeletePlanResponse:
    result = ProductionService().delete_plan(plan_id)
    if not result.get("deleted"):
        raise HTTPException(status_code=404, detail="Plan not found")
    return DeletePlanResponse(**result)


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


@router.get("/day-occupancy", response_model=DayOccupancyResponse)
def get_day_occupancy(
    exclude_plan_id: str | None = None,
    _user: dict = Depends(require_roles("admin", "production")),
) -> DayOccupancyResponse:
    result = ProductionService().get_day_occupancy(exclude_plan_id=exclude_plan_id)
    return DayOccupancyResponse(**result)


@router.get("/kp-candidates", response_model=KpCandidatesResponse)
def get_kp_candidates(
    _user: dict = Depends(require_roles("admin", "production")),
) -> KpCandidatesResponse:
    result = ProductionService().list_kp_candidates()
    return KpCandidatesResponse(**result)


@router.get("/days/{target_date}", response_model=DayViewDetailResponse)
def get_day_view(
    target_date: str,
    _user: dict = Depends(require_roles("admin", "production")),
) -> DayViewDetailResponse:
    data = ProductionService().get_day_view_detailed(target_date)
    if not data:
        raise HTTPException(status_code=404, detail="Day not found")
    return DayViewDetailResponse(**data)


@router.post("/days/{target_date}/complete")
def complete_day(
    target_date: str,
    payload: CompleteProductionDayRequest,
    _user: dict = Depends(require_roles("admin", "production")),
) -> dict:
    return ProductionService().complete_day(plan_id=payload.plan_id, target_date=target_date)


@router.get("/days/{target_date}/documents/schema")
async def download_day_schema(
    target_date: str,
    background_tasks: BackgroundTasks,
    _user: dict = Depends(require_roles("admin", "production")),
) -> FileResponse:
    try:
        pdf_path, cleanup_dir = await generate_day_schema(target_date)
    except DayDocumentsError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except Exception as exc:
        logger.exception("[production/day-schema] ошибка: %s", exc)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось сформировать схему дорожек.",
        ) from exc
    background_tasks.add_task(make_cleanup_callback(cleanup_dir))
    return FileResponse(
        path=str(pdf_path),
        media_type="application/pdf",
        filename=f"Схема_{target_date}.pdf",
    )


@router.get("/days/{target_date}/documents/breakdown")
async def download_day_breakdown(
    target_date: str,
    background_tasks: BackgroundTasks,
    _user: dict = Depends(require_roles("admin", "production")),
) -> FileResponse:
    try:
        xlsx_path, cleanup_dir = await generate_day_breakdown(target_date)
    except DayDocumentsError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except Exception as exc:
        logger.exception("[production/day-breakdown] ошибка: %s", exc)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось сформировать детальную разбивку.",
        ) from exc
    background_tasks.add_task(make_cleanup_callback(cleanup_dir))
    return FileResponse(
        path=str(xlsx_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"Детальная_разбивка_{target_date}.xlsx",
    )


@router.get("/days/{target_date}/documents/formovka")
async def download_day_formovka(
    target_date: str,
    background_tasks: BackgroundTasks,
    _user: dict = Depends(require_roles("admin", "production")),
) -> FileResponse:
    try:
        zip_path, cleanup_dir = await generate_day_formovka(target_date)
    except DayDocumentsError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except Exception as exc:
        logger.exception("[production/day-formovka] ошибка: %s", exc)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось сформировать файлы формовки.",
        ) from exc
    background_tasks.add_task(make_cleanup_callback(cleanup_dir))
    return FileResponse(
        path=str(zip_path),
        media_type="application/zip",
        filename=f"Формовка_{target_date}.zip",
    )


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
