"""Стол прайсов: preview / apply / status."""

from __future__ import annotations

from collections.abc import Callable
from typing import TypeVar

from fastapi import APIRouter, Depends, File, Form, UploadFile

from app.core.http_errors import raise_client_error, raise_unexpected_server_error
from app.dependencies.auth import REQUIRE_PRICES
from app.dependencies.services import get_pile_cost_service, get_price_desk_service
from app.schemas.pile_cost import (
    PileCostApplyRequest,
    PileCostInputsBody,
    PileCostPreview,
    PileCostSaved,
    PileCostState,
)
from app.schemas.price_desk import PriceDeskPreview, PriceDeskStatus
from app.services.commercial_upload_validation import read_upload_file_capped
from app.services.pile_cost_service import PileCostConflict, PileCostNotFound, PileCostService
from app.services.price_desk_service import PriceDeskError, PriceDeskService

router = APIRouter(prefix="/prices", tags=["prices"])


@router.get("/status", response_model=PriceDeskStatus)
def prices_status(
    _user: dict = Depends(REQUIRE_PRICES),
    service: PriceDeskService = Depends(get_price_desk_service),
) -> PriceDeskStatus:
    return service.status()


@router.post("/preview", response_model=PriceDeskPreview)
async def prices_preview(
    file: UploadFile = File(..., description="Заводской прайс (.xls / .xlsx)"),
    _user: dict = Depends(REQUIRE_PRICES),
    service: PriceDeskService = Depends(get_price_desk_service),
) -> PriceDeskPreview:
    file_bytes = await read_upload_file_capped(file)
    try:
        return service.preview(file_bytes, file.filename)
    except PriceDeskError as exc:
        raise_client_error(
            exc,
            status_code=exc.status_code,
            detail=exc.message,
            where="prices.preview",
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="prices.preview")


@router.post("/apply", response_model=PriceDeskPreview)
async def prices_apply(
    file: UploadFile = File(..., description="Тот же файл, что в превью"),
    file_sha256: str | None = Form(default=None),
    _user: dict = Depends(REQUIRE_PRICES),
    service: PriceDeskService = Depends(get_price_desk_service),
) -> PriceDeskPreview:
    file_bytes = await read_upload_file_capped(file)
    try:
        return service.apply(file_bytes, file.filename, file_sha256)
    except PriceDeskError as exc:
        raise_client_error(
            exc,
            status_code=exc.status_code,
            detail=exc.message,
            where="prices.apply",
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="prices.apply")


_T = TypeVar("_T")


def _pile_cost_call(action: Callable[[], _T], *, where: str) -> _T:
    try:
        return action()
    except PileCostConflict as exc:
        raise_client_error(exc, status_code=exc.status_code, detail=exc.message, where=where)
    except PileCostNotFound as exc:
        raise_client_error(exc, status_code=exc.status_code, detail=exc.message, where=where)
    except ValueError as exc:
        raise_client_error(exc, status_code=422, detail=str(exc), where=where)
    except Exception as exc:
        raise_unexpected_server_error(exc, where=where)


@router.get("/pile-cost", response_model=PileCostState)
def pile_cost_get(
    mark: str | None = None,
    _user: dict = Depends(REQUIRE_PRICES),
    service: PileCostService = Depends(get_pile_cost_service),
) -> PileCostState:
    return _pile_cost_call(lambda: service.view(mark), where="prices.pile_cost.get")


@router.put("/pile-cost", response_model=PileCostSaved)
def pile_cost_put(
    body: PileCostInputsBody,
    _user: dict = Depends(REQUIRE_PRICES),
    service: PileCostService = Depends(get_pile_cost_service),
) -> PileCostSaved:
    return _pile_cost_call(lambda: service.save(body), where="prices.pile_cost.put")


@router.post("/pile-cost/preview", response_model=PileCostPreview)
def pile_cost_preview(
    _user: dict = Depends(REQUIRE_PRICES),
    service: PileCostService = Depends(get_pile_cost_service),
) -> PileCostPreview:
    return _pile_cost_call(service.preview, where="prices.pile_cost.preview")


@router.post("/pile-cost/apply", response_model=PileCostPreview)
def pile_cost_apply(
    body: PileCostApplyRequest | None = None,
    _user: dict = Depends(REQUIRE_PRICES),
    service: PileCostService = Depends(get_pile_cost_service),
) -> PileCostPreview:
    digest = None if body is None else body.inputs_sha256
    return _pile_cost_call(lambda: service.apply(digest), where="prices.pile_cost.apply")
