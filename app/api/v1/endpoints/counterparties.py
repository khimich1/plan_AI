from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends, Query
from fastapi.responses import JSONResponse

from app.dependencies.auth import require_roles
from app.dependencies.services import get_counterparties_service
from app.schemas.counterparties import (
    COUNTERPARTY_SEARCH_Q_MAX_LENGTH,
    CounterpartyCreateRequest,
    CounterpartyCreateResponse,
    CounterpartySearchResponse,
    CounterpartyShort,
)
from app.services.counterparties_service import (
    CounterpartiesService,
    DuplicateCodeError,
)

router = APIRouter(prefix="/counterparties", tags=["counterparties"])


def _to_short(row: dict[str, Any]) -> CounterpartyShort:
    return CounterpartyShort(
        id=int(row["id"]),
        code_1c=str(row["code_1c"]),
        name=str(row["name"]),
        inn=row.get("inn"),
        kpp=row.get("kpp"),
    )


@router.get("/search", response_model=CounterpartySearchResponse)
def search_counterparties(
    q: str = Query(..., min_length=2, max_length=COUNTERPARTY_SEARCH_Q_MAX_LENGTH),
    limit: int = Query(10, ge=1, le=50),
    clients_only: bool = Query(True),
    user: dict = Depends(require_roles("admin", "manager")),
    service: CounterpartiesService = Depends(get_counterparties_service),
) -> CounterpartySearchResponse:
    if user.get("role") != "admin":
        clients_only = True
    items = [_to_short(row) for row in service.search(q, limit=limit, clients_only=clients_only)]
    return CounterpartySearchResponse(items=items, count=len(items))


@router.post("", response_model=CounterpartyCreateResponse, status_code=201)
def create_counterparty(
    payload: CounterpartyCreateRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
    service: CounterpartiesService = Depends(get_counterparties_service),
) -> CounterpartyCreateResponse | JSONResponse:
    try:
        result = service.create(
            name=payload.name,
            code_1c=payload.code_1c,
            inn=payload.inn,
            kpp=payload.kpp,
            is_client=payload.is_client,
        )
    except DuplicateCodeError as exc:
        return JSONResponse(
            status_code=409,
            content={
                "detail": str(exc),
                "existing": _to_short(exc.existing).model_dump(),
            },
        )
    except ValueError as exc:
        return JSONResponse(status_code=422, content={"detail": str(exc)})
    return CounterpartyCreateResponse(item=_to_short(result.item), warning=result.warning)
