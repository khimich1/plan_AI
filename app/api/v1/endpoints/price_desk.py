"""Стол прайсов: preview / apply / status."""

from __future__ import annotations

from fastapi import APIRouter, Depends, File, Form, UploadFile

from app.core.http_errors import raise_client_error, raise_unexpected_server_error
from app.dependencies.auth import REQUIRE_PRICES
from app.dependencies.services import get_price_desk_service
from app.schemas.price_desk import PriceDeskPreview, PriceDeskStatus
from app.services.commercial_upload_validation import read_upload_file_capped
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
