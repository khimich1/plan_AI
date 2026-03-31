from __future__ import annotations

from fastapi import APIRouter, Depends

from app.dependencies.auth import require_roles
from app.schemas.commercial import CommercialParseRequest, CommercialPreviewRequest
from app.services.commercial_service import CommercialService
from app.services.draft_store import DraftStore

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

