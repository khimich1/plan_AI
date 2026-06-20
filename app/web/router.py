from __future__ import annotations

from html import escape
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile, status
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse

from app.core.settings import get_settings
from app.web.legacy_deprecation import (
    SPA_ARCHIVE,
    SPA_LOGIN,
    SPA_NEW,
    SPA_PRODUCTION,
    default_spa_home_for_role,
    deprecated_redirect,
    mark_legacy_response,
    spa_draft_url,
    spa_new_with_notice,
    spa_new_with_error,
)
from app.dependencies.auth import REQUIRE_ADMIN_OR_MANAGER, get_current_user, require_roles
from app.dependencies.commercial_draft import check_draft_ownership
from app.dependencies.plate_context import get_plate_order_context
from app.repositories.auth_repository import AuthRepository
from app.security.csrf import clear_csrf_cookie, generate_csrf_token, set_csrf_cookie
from app.security.login_rate_limit import check_login_rate_limit, resolve_client_ip
from app.security.session import clear_session_cookie, create_session_token, set_session_cookie
from app.services.commercial_service import CommercialService
from app.services.commercial_upload_validation import prepare_commercial_ocr_upload
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.offers_service import OffersService
from app.services.production_service import ProductionService
from core.exceptions import PlateParseError

router = APIRouter(include_in_schema=False)


def _page(title: str, body: str) -> HTMLResponse:
    html = f"""
    <!doctype html>
    <html lang="ru">
      <head>
        <meta charset="utf-8">
        <title>{escape(title)}</title>
        <style>
          body {{ font-family: Arial, sans-serif; margin: 2rem; }}
          nav a {{ margin-right: 1rem; }}
          table {{ border-collapse: collapse; width: 100%; margin-top: 1rem; }}
          th, td {{ border: 1px solid #ccc; padding: 0.5rem; text-align: left; }}
          .card {{ border: 1px solid #ddd; padding: 1rem; margin-bottom: 1rem; border-radius: 8px; }}
        </style>
      </head>
      <body>{body}</body>
    </html>
    """
    return HTMLResponse(html)


def _render_frontend_shell() -> HTMLResponse:
    settings = get_settings()
    index_path = settings.frontend_dist_dir / "index.html"
    if not index_path.exists():
        body = """
        <h1>Frontend build РЅРµ РЅР°Р№РґРµРЅ</h1>
        <p>РЎРѕР±РµСЂРёС‚Рµ React-РїСЂРёР»РѕР¶РµРЅРёРµ РІ РґРёСЂРµРєС‚РѕСЂРёСЋ <code>frontend/dist</code>, Р·Р°С‚РµРј РѕС‚РєСЂРѕР№С‚Рµ СЃС‚СЂР°РЅРёС†Сѓ СЃРЅРѕРІР°.</p>
        <p>Р”Р»СЏ Р»РѕРєР°Р»СЊРЅРѕР№ СЂР°Р·СЂР°Р±РѕС‚РєРё РёСЃРїРѕР»СЊР·СѓР№С‚Рµ Vite dev server РёР· РґРёСЂРµРєС‚РѕСЂРёРё <code>frontend/</code>.</p>
        """
        return _page("Frontend build missing", body)
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


def _resolve_frontend_asset(asset_path: str) -> Path | None:
    settings = get_settings()
    assets_dir = (settings.frontend_dist_dir / "assets").resolve()
    candidate = (assets_dir / asset_path).resolve()
    if assets_dir not in candidate.parents or not candidate.exists():
        return None
    return candidate






def _legacy_new_offer_error_response(request: Request, message: str) -> RedirectResponse | JSONResponse:
    accept = (request.headers.get("accept") or "").lower()
    successor = spa_new_with_error(message)
    if "application/json" in accept and "text/html" not in accept:
        response = JSONResponse({"detail": message}, status_code=400)
        return mark_legacy_response(response, legacy_path="/web/offers/new", successor=successor)
    response = RedirectResponse(successor, status_code=303)
    return mark_legacy_response(response, legacy_path="/web/offers/new", successor=successor)


def _redact_saved_offer_if_forbidden(draft: dict, user: dict) -> dict:
    saved_offer = draft.get("saved_offer")
    if not saved_offer or saved_offer.get("kp_id") is None:
        return draft
    try:
        OffersService().get_offer(int(saved_offer["kp_id"]), user=user)
    except HTTPException:
        draft = dict(draft)
        draft["saved_offer"] = None
    return draft



@router.get("/web/login")
def login_page(request: Request) -> RedirectResponse:
    _ = request
    return deprecated_redirect(SPA_LOGIN, legacy_path="/web/login")


@router.post("/web/login")
def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
) -> RedirectResponse:
    check_login_rate_limit(resolve_client_ip(request))
    repository = AuthRepository()
    user = repository.authenticate(username, password)
    if not user:
        response = RedirectResponse(f"{SPA_LOGIN}?error=РќРµРІРµСЂРЅС‹Рµ+РґР°РЅРЅС‹Рµ", status_code=303)
        return mark_legacy_response(response, legacy_path="/web/login", successor=SPA_LOGIN)
    token = create_session_token({"id": user["id"], "username": user["username"], "role": user["role"]})
    home = default_spa_home_for_role(user["role"])
    response = RedirectResponse(home, status_code=303)
    set_session_cookie(response, token)
    set_csrf_cookie(response, generate_csrf_token())
    return mark_legacy_response(response, legacy_path="/web/login", successor=home)


@router.get("/web/logout")
def web_logout() -> RedirectResponse:
    response = RedirectResponse(SPA_LOGIN, status_code=303)
    clear_session_cookie(response)
    clear_csrf_cookie(response)
    return mark_legacy_response(response, legacy_path="/web/logout", successor=SPA_LOGIN)


@router.get("/web")
def dashboard(user: dict = Depends(get_current_user)) -> RedirectResponse:
    target = default_spa_home_for_role(user.get("role"))
    return deprecated_redirect(target, legacy_path="/web")



@router.get("/web/managers")
def managers_page(
    user: dict = Depends(require_roles("admin", "manager", "production")),
) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_ARCHIVE, legacy_path="/web/managers")



@router.get("/web/offers")
def offers_page(
    user: dict = Depends(require_roles("admin", "manager", "production")),
) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_ARCHIVE, legacy_path="/web/offers")



@router.get("/web/offers/new")
def new_offer_page(user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER)) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_NEW, legacy_path="/web/offers/new")


@router.post("/web/offers/new", response_model=None)
async def new_offer_submit(
    request: Request,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    text: str = Form(default=""),
    manager_id: str = Form(default=""),
    client_name: str = Form(default=""),
    discount_percent: str = Form(default="0"),
    delivery_conditions: str = Form(default=""),
    payment_conditions: str = Form(default=""),
    image: UploadFile | None = File(default=None),
) -> RedirectResponse | JSONResponse:
    try:
        parsed_manager_id = int(manager_id)
    except ValueError:
        return _legacy_new_offer_error_response(request, "Выберите менеджера.")

    try:
        parsed_discount = float((discount_percent or "0").replace(",", "."))
    except ValueError:
        return _legacy_new_offer_error_response(request, "Скидка должна быть числом.")

    try:
        image_bytes, image_name = await prepare_commercial_ocr_upload(
            image=image,
            user_id=int(user["id"]),
        )
    except HTTPException as exc:
        msg = exc.detail if isinstance(exc.detail, str) else "РћС€РёР±РєР° Р·Р°РіСЂСѓР·РєРё С„Р°Р№Р»Р°."
        return _legacy_new_offer_error_response(request, msg)

    workflow = CommercialWorkflowService()
    try:
        draft = await workflow.create_draft_from_form(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_name,
            manager_id=parsed_manager_id,
            client_name=client_name,
            discount_percent=parsed_discount,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            owner_user_id=int(user["id"]),
        )
    except (PlateParseError, ValueError) as exc:
        return _legacy_new_offer_error_response(request, str(exc))
    response = RedirectResponse(spa_draft_url(draft["draft_id"]), status_code=303)
    successor = spa_draft_url(draft["draft_id"])
    return mark_legacy_response(response, legacy_path="/web/offers/new", successor=successor)


@router.get("/web/offers/drafts/{draft_id}")
def offer_draft_page(
    draft_id: str,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> RedirectResponse:
    legacy_path = f"/web/offers/drafts/{draft_id}"
    try:
        check_draft_ownership(draft_id, user)
    except HTTPException:
        return deprecated_redirect(
            spa_new_with_notice("Р§РµСЂРЅРѕРІРёРє РЅРµ РЅР°Р№РґРµРЅ РёР»Рё РЅРµРґРѕСЃС‚СѓРїРµРЅ."),
            legacy_path=legacy_path,
        )
    workflow = CommercialWorkflowService()
    try:
        workflow.get_draft_details(draft_id)
    except FileNotFoundError:
        return deprecated_redirect(
            spa_new_with_notice("Р§РµСЂРЅРѕРІРёРє РЅРµ РЅР°Р№РґРµРЅ РёР»Рё РЅРµРґРѕСЃС‚СѓРїРµРЅ."),
            legacy_path=legacy_path,
        )
    return deprecated_redirect(spa_draft_url(draft_id), legacy_path=legacy_path)


@router.post("/web/offers/drafts/{draft_id}/generate-files")
def generate_offer_draft_files(
    draft_id: str,
    request: Request,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> RedirectResponse:
    legacy_path = f"/web/offers/drafts/{draft_id}/generate-files"
    successor = spa_draft_url(draft_id)
    try:
        check_draft_ownership(draft_id, user)
    except HTTPException:
        response = RedirectResponse(SPA_ARCHIVE, status_code=303)
        return mark_legacy_response(response, legacy_path=legacy_path, successor=SPA_ARCHIVE)
    workflow = CommercialWorkflowService()
    try:
        workflow.generate_files(
            draft_id,
            plate_order_ctx=get_plate_order_context(request),
        )
    except (FileNotFoundError, ValueError):
        response = RedirectResponse(
            spa_new_with_notice("РќРµ СѓРґР°Р»РѕСЃСЊ СЃРіРµРЅРµСЂРёСЂРѕРІР°С‚СЊ С„Р°Р№Р»С‹ РґР»СЏ С‡РµСЂРЅРѕРІРёРєР°."),
            status_code=303,
        )
        return mark_legacy_response(response, legacy_path=legacy_path, successor=successor)
    response = RedirectResponse(successor, status_code=303)
    return mark_legacy_response(response, legacy_path=legacy_path, successor=successor)


@router.post("/web/offers/drafts/{draft_id}/save")
def save_offer_draft(
    draft_id: str,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> RedirectResponse:
    legacy_path = f"/web/offers/drafts/{draft_id}/save"
    successor = spa_draft_url(draft_id)
    try:
        check_draft_ownership(draft_id, user)
    except HTTPException:
        response = RedirectResponse(SPA_ARCHIVE, status_code=303)
        return mark_legacy_response(response, legacy_path=legacy_path, successor=SPA_ARCHIVE)
    workflow = CommercialWorkflowService()
    try:
        workflow.save_offer(draft_id)
    except FileNotFoundError:
        response = RedirectResponse(
            spa_new_with_notice("РќРµ СѓРґР°Р»РѕСЃСЊ СЃРѕС…СЂР°РЅРёС‚СЊ С‡РµСЂРЅРѕРІРёРє РІ Р±Р°Р·Сѓ."),
            status_code=303,
        )
        return mark_legacy_response(response, legacy_path=legacy_path, successor=successor)
    response = RedirectResponse(successor, status_code=303)
    return mark_legacy_response(response, legacy_path=legacy_path, successor=successor)


@router.get("/web/production")
def production_page(user: dict = Depends(require_roles("admin", "production"))) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_PRODUCTION, legacy_path="/web/production")



@router.get("/commercial-offer/login", response_class=HTMLResponse)
def commercial_offer_login_spa() -> HTMLResponse:
    return _render_frontend_shell()


@router.get("/commercial-offer/archive", response_class=HTMLResponse)
def commercial_offer_archive_spa(
    user: dict = Depends(require_roles("admin", "manager", "production")),
) -> HTMLResponse:
    _ = user
    return _render_frontend_shell()


@router.get("/commercial-offer/production", response_class=HTMLResponse)
def commercial_offer_production_spa(
    user: dict = Depends(require_roles("admin", "production")),
) -> HTMLResponse:
    _ = user
    return _render_frontend_shell()


@router.get("/commercial-offer")
def commercial_offer_root(user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER)) -> RedirectResponse:
    _ = user
    return RedirectResponse("/commercial-offer/new", status_code=303)


@router.get("/commercial-offer/new", response_class=HTMLResponse)
def commercial_offer_spa(user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER)) -> HTMLResponse:
    _ = user
    return _render_frontend_shell()


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
    target = _resolve_frontend_asset(asset_path)
    if target is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Frontend asset not found.")
    return FileResponse(path=target)



