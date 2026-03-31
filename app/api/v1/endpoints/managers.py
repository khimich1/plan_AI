from __future__ import annotations

from fastapi import APIRouter, Depends

from app.dependencies.auth import require_roles
from app.services.commercial_service import CommercialService

router = APIRouter(prefix="/managers", tags=["managers"])


@router.get("")
def list_managers(
    _user: dict = Depends(require_roles("admin", "manager", "production")),
) -> dict:
    service = CommercialService()
    managers = service.list_managers()
    return {"items": managers, "count": len(managers)}

