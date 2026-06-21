from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse

from app.dependencies.auth import REQUIRE_ADMIN_OR_MANAGER, require_roles
from app.web.legacy_deprecation import spa_draft_url
from app.web.shell import render_frontend_shell, resolve_frontend_asset

router = APIRouter(include_in_schema=False)


@router.get("/commercial-offer/login", response_class=HTMLResponse)
def commercial_offer_login_spa() -> HTMLResponse:
    return render_frontend_shell()


@router.get("/commercial-offer/archive", response_class=HTMLResponse)
def commercial_offer_archive_spa(
    user: dict = Depends(require_roles("admin", "manager", "production")),
) -> HTMLResponse:
    _ = user
    return render_frontend_shell()


@router.get("/commercial-offer/production", response_class=HTMLResponse)
def commercial_offer_production_spa(
    user: dict = Depends(require_roles("admin", "production")),
) -> HTMLResponse:
    _ = user
    return render_frontend_shell()


@router.get("/commercial-offer")
def commercial_offer_root(user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER)) -> RedirectResponse:
    _ = user
    return RedirectResponse("/commercial-offer/new", status_code=303)


@router.get("/commercial-offer/new", response_class=HTMLResponse)
def commercial_offer_spa(user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER)) -> HTMLResponse:
    _ = user
    return render_frontend_shell()


@router.get("/commercial-offer/drafts/{draft_id}")
def commercial_offer_legacy_draft_stub(
    draft_id: str,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> RedirectResponse:
    _ = user
    return RedirectResponse(spa_draft_url(draft_id), status_code=303)


@router.get("/commercial-offer/assets/{asset_path:path}")
def commercial_offer_assets(
    asset_path: str,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> FileResponse:
    _ = user
    target = resolve_frontend_asset(asset_path)
    if target is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Frontend asset not found.")
    return FileResponse(path=target)
