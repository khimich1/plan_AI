from __future__ import annotations

import os
import tempfile
from datetime import datetime

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import Response

from app.core.settings import get_settings
from app.dependencies.auth import require_roles
from app.schemas.commercial import (
    CommercialParseRequest,
    CommercialPreviewCheckXlsxRequest,
    CommercialPreviewRequest,
    CommercialPreviewXlsxRequest,
    CommercialRecognizeScreenResponse,
)
from app.services.commercial_service import CommercialService
from app.services.draft_store import DraftStore
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.ocr_gpt import GPT_AVAILABLE, recognize_text_smart
from core.plate_text_normalizer import normalize_order_text
from core.plates_preview_xlsx import build_plates_reconciliation_preview_xlsx
from core.reconciliation_xlsx import split_plate_text_lines

router = APIRouter(prefix="/commercial", tags=["commercial"])


@router.post("/parse")
def parse_commercial_text(
    payload: CommercialParseRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> dict:
    service = CommercialService()
    result = service.parse(payload.text)
    return {
        "order": result.order.to_dict(),
        "normalized_text": result.normalized_text,
        "normalized_lines": result.normalized_lines,
        "unparsed_lines": result.unparsed_lines,
        "warnings": result.warnings,
        "wide_plate_lines": result.wide_plate_lines,
        "diagnostics": result.diagnostics,
    }


@router.post("/generate-preview")
def generate_preview(
    payload: CommercialPreviewRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> dict:
    service = CommercialService()
    preview = service.generate_preview(text=payload.text)
    draft_id = DraftStore().save_preview(
        order=preview.parse_result.order,
        optimization_context=preview.optimization_context,
        order_data=preview.order_data,
        metadata={
            "normalized_text": preview.parse_result.normalized_text,
            "warnings": preview.parse_result.warnings,
        },
    )
    return {
        "draft_id": draft_id,
        "order": preview.parse_result.order.to_dict(),
        "unparsed_lines": preview.parse_result.unparsed_lines,
        "warnings": preview.parse_result.warnings,
        "optimization": {
            "total_plates": preview.optimization_context.total_plates,
            "total_cost": preview.optimization_context.total_cost,
        },
        "order_data": preview.order_data,
        "price_rows_count": len(preview.price_rows),
        "breakdown_tables_count": len(preview.breakdown_tables),
        "total_sum": preview.total_sum,
    }


@router.get("/drafts/{draft_id}")
def get_preview_draft(
    draft_id: str,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> dict:
    payload = DraftStore().load_preview(draft_id)
    if not payload:
        return {"draft_id": draft_id, "found": False}
    return {
        "draft_id": draft_id,
        "found": True,
        "order": payload["order"].to_dict(),
        "optimization": payload["optimization_context"].optimization_result,
        "order_data": payload["order_data"],
        "metadata": payload.get("metadata", {}),
    }


@router.post("/drafts/{draft_id}/xlsx")
def download_preview_xlsx(
    draft_id: str,
    payload: CommercialPreviewXlsxRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> Response:
    draft_payload = DraftStore().load_preview(draft_id)
    if not draft_payload:
        raise HTTPException(status_code=404, detail="Preview draft not found")
    order_data = draft_payload.get("order_data") or []
    if not order_data:
        raise HTTPException(status_code=400, detail="Preview draft has no order data")

    offer_date = datetime.now().strftime("%d.%m.%Y")
    offer_number = f"preview-{draft_id[:8]}"
    xlsx_buffer = generate_commercial_offer_xlsx(
        order_data=order_data,
        offer_number=offer_number,
        offer_date=offer_date,
        customer_name=payload.customer_name,
        manager_name=payload.manager_name,
        manager_phone=payload.manager_phone or None,
        manager_email=payload.manager_email or None,
        discount_percent=payload.discount_percent,
        delivery_conditions=payload.delivery_conditions or None,
        payment_conditions=payload.payment_conditions or None,
        kp_db_id=None,
    )

    filename = f"KP_preview_{draft_id[:8]}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=xlsx_buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@router.post("/recognize-screen", response_model=CommercialRecognizeScreenResponse)
async def recognize_screen(
    image: UploadFile = File(...),
    _user: dict = Depends(require_roles("admin", "manager")),
) -> CommercialRecognizeScreenResponse:
    if not image.filename:
        raise HTTPException(status_code=400, detail="Image file is required")
    content_type = image.content_type or ""
    if not content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="Only image files are supported")

    suffix = os.path.splitext(image.filename)[1] or ".jpg"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
        temp_path = temp_file.name
        temp_file.write(await image.read())

    try:
        settings = get_settings()
        recognition_mode = (settings.ocr_recognition_mode or "full_gpt").strip().lower()
        if recognition_mode not in {"full_gpt", "hybrid"}:
            recognition_mode = "full_gpt"
        result = await recognize_text_smart(
            temp_path,
            force_gpt=(recognition_mode == "full_gpt"),
            show_cost=False,
            mode=recognition_mode,
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"OCR failed: {exc}") from exc
    finally:
        try:
            os.unlink(temp_path)
        except OSError:
            pass

    if not result or not result.get("text"):
        if not GPT_AVAILABLE:
            raise HTTPException(
                status_code=500,
                detail="OCR недоступен: установите зависимость `openai` в активное venv.",
            )
        raise HTTPException(status_code=500, detail="OCR failed: empty recognition result")
    recognized_text = (result.get("text") or "").strip()
    normalized_text = normalize_order_text(recognized_text)
    lines = split_plate_text_lines(normalized_text)
    warnings: list[str] = []
    if not lines:
        warnings.append("Распознавание выполнено, но строки плит не выделены.")

    return CommercialRecognizeScreenResponse(
        recognized_text=recognized_text,
        normalized_text=normalized_text,
        lines=lines,
        warnings=warnings,
        method=str(result.get("method") or ""),
        confidence=float(result.get("confidence") or 0.0),
    )


@router.post("/preview-check-xlsx")
def download_preview_check_xlsx(
    payload: CommercialPreviewCheckXlsxRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> Response:
    recognized_text = (payload.recognized_text or "").strip()
    plates_text = (payload.plates_text or "").strip()
    if not plates_text:
        raise HTTPException(status_code=400, detail="plates_text is required")
    source_for_parse = recognized_text or plates_text
    initial_user_lines = split_plate_text_lines(plates_text)
    if not initial_user_lines:
        raise HTTPException(status_code=400, detail="No source lines found in plates_text")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            tmp_path = temp_file.name
        build_plates_reconciliation_preview_xlsx(
            tmp_path,
            plates_text=source_for_parse,
            initial_user_plate_lines=initial_user_lines,
            forced_wide_line_indexes=[],
        )
        with open(tmp_path, "rb") as file:
            content = file.read()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to generate preview check XLSX: {exc}") from exc
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    filename = f"preview_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

