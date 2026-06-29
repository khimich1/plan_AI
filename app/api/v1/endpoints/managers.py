from __future__ import annotations

from fastapi import APIRouter, Depends

from app.dependencies.auth import require_roles
from app.dependencies.services import get_commercial_service
from app.schemas.managers import ManagerListResponse
from app.services.commercial_service import CommercialService

router = APIRouter(prefix="/managers", tags=["managers"])


@router.get("", response_model=ManagerListResponse)
def list_managers(
    _user: dict = Depends(require_roles("admin", "manager", "production")),
    service: CommercialService = Depends(get_commercial_service),
) -> ManagerListResponse:
    managers = service.list_managers()
    return ManagerListResponse.model_validate({"items": managers, "count": len(managers)})
