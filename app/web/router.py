from __future__ import annotations

from html import escape

from fastapi import APIRouter, Depends, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse

from app.dependencies.auth import get_current_user, require_roles
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from app.services.commercial_service import CommercialService
from app.services.commercial_workflow_service import CommercialWorkflowService
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


def _nav(user: dict) -> str:
    return (
        f"<nav>"
        f"<strong>{escape(user['username'])}</strong> ({escape(user['role'])}) "
        f'<a href="/web">Главная</a>'
        f'<a href="/web/managers">Менеджеры</a>'
        f'<a href="/web/offers">КП</a>'
        f'<a href="/web/production">Производство</a>'
        f'<a href="/web/login">Выйти</a>'
        f"</nav><hr>"
    )


def _render_offer_form(
    *,
    user: dict,
    managers: list[dict],
    error: str = "",
    values: dict[str, str] | None = None,
) -> HTMLResponse:
    values = values or {}
    options = ['<option value="">Выберите менеджера</option>']
    selected_manager = values.get("manager_id", "")
    for manager in managers:
        manager_id = str(manager.get("id", ""))
        selected_attr = " selected" if manager_id == selected_manager else ""
        options.append(
            f'<option value="{escape(manager_id)}"{selected_attr}>{escape(manager.get("fio", ""))}</option>'
        )

    body = _nav(user) + "<h1>Новое коммерческое предложение</h1>"
    if error:
        body += f'<p style="color: red;">{escape(error)}</p>'
    body += f"""
    <div class="card">
      <form method="post" action="/web/offers/new" enctype="multipart/form-data">
        <label>Менеджер<br><select name="manager_id" required>{"".join(options)}</select></label><br><br>
        <label>Клиент<br><input type="text" name="client_name" value="{escape(values.get("client_name", ""))}" required></label><br><br>
        <label>Скидка, %<br><input type="number" name="discount_percent" min="0" max="100" step="0.01" value="{escape(values.get("discount_percent", "0"))}"></label><br><br>
        <label>Условия поставки<br><textarea name="delivery_conditions" rows="3" cols="80">{escape(values.get("delivery_conditions", ""))}</textarea></label><br><br>
        <label>Условия оплаты<br><textarea name="payment_conditions" rows="3" cols="80">{escape(values.get("payment_conditions", ""))}</textarea></label><br><br>
        <label>Текст плит<br><textarea name="text" rows="12" cols="100" placeholder="ПБ 78-12-8п 2&#10;ПБ 66-3-8п 4">{escape(values.get("text", ""))}</textarea></label><br><br>
        <label>Или изображение / скан<br><input type="file" name="image" accept="image/*"></label><br><br>
        <button type="submit">Создать черновик КП</button>
      </form>
    </div>
    """
    return _page("New Offer", body)


def _render_offer_preview(*, user: dict, draft: dict) -> HTMLResponse:
    metadata = draft.get("metadata", {})
    totals = draft.get("totals", {})
    escaped_draft_id = escape(str(draft.get("draft_id", "")))
    order_rows = "".join(
        f"<tr>"
        f"<td>{escape(item.get('name', ''))}</td>"
        f"<td>{item.get('qty', 0)}</td>"
        f"<td>{item.get('length_m', 0)}</td>"
        f"<td>{item.get('width_m', 0)}</td>"
        f"<td>{item.get('unit_price', 0)}</td>"
        f"</tr>"
        for item in draft.get("order_data", [])
    )
    warnings = "".join(f"<li>{escape(str(item))}</li>" for item in metadata.get("warnings", []))
    unparsed = "".join(f"<li>{escape(str(item))}</li>" for item in metadata.get("unparsed_lines", []))
    files = "".join(
        f'<li><a href="{escape(file["download_url"])}">{escape(file["display_name"])}</a></li>'
        for file in draft.get("files", [])
    )
    saved_offer = draft.get("saved_offer")

    body = _nav(user) + "<h1>Черновик коммерческого предложения</h1>"
    body += f"""
    <div class="card">
      <strong>Draft ID:</strong> {escaped_draft_id}<br>
      <strong>Клиент:</strong> {escape(metadata.get("client_name", ""))}<br>
      <strong>Менеджер:</strong> {escape(metadata.get("manager_name", ""))}<br>
      <strong>Скидка:</strong> {metadata.get("discount_percent", 0)}%<br>
      <strong>Источник:</strong> {escape(metadata.get("source_type", ""))}<br>
      <strong>Распознано плит:</strong> {sum(int(item.get("qty", 0) or 0) for item in draft.get("order_data", []))}<br>
      <strong>Итого с НДС:</strong> {totals.get("total_with_vat", 0)}<br>
      <strong>Оптимизация:</strong> {draft.get("optimization", {}).get("total_plates", 0)} плит, {draft.get("optimization", {}).get("total_cost", 0)} ₽
    </div>
    <div class="card">
      <form method="post" action="/web/offers/drafts/{escaped_draft_id}/generate-files">
        <button type="submit">Сгенерировать файлы</button>
      </form>
      <br>
      <form method="post" action="/web/offers/drafts/{escaped_draft_id}/save">
        <button type="submit">Сохранить КП в базу</button>
      </form>
    </div>
    """
    if draft.get("files"):
        body += f'<div class="card"><h2>Файлы</h2><ul>{files}</ul></div>'
    if saved_offer:
        body += (
            '<div class="card">'
            f"<strong>Сохранено в БД:</strong> КП #{saved_offer.get('kp_id')} ({escape(saved_offer.get('status', ''))})"
            "</div>"
        )
    if warnings:
        body += f'<div class="card"><h2>Предупреждения</h2><ul>{warnings}</ul></div>'
    if unparsed:
        body += f'<div class="card"><h2>Нераспознанные строки</h2><ul>{unparsed}</ul></div>'
    body += (
        '<div class="card"><h2>Позиции</h2>'
        f"<table><tr><th>Наименование</th><th>Кол-во</th><th>Длина, м</th><th>Ширина, м</th><th>Цена</th></tr>{order_rows}</table>"
        "</div>"
    )
    return _page("Offer Draft", body)


@router.get("/web/login", response_class=HTMLResponse)
def login_page(request: Request) -> HTMLResponse:
    error = request.query_params.get("error")
    body = """
    <h1>Вход в систему</h1>
    <form method="post" action="/web/login">
      <label>Логин <input type="text" name="username"></label><br><br>
      <label>Пароль <input type="password" name="password"></label><br><br>
      <button type="submit">Войти</button>
    </form>
    """
    if error:
        body = f'<p style="color: red;">{escape(error)}</p>' + body
    return _page("Login", body)


@router.post("/web/login")
def login_submit(username: str = Form(...), password: str = Form(...)) -> RedirectResponse:
    repository = AuthRepository()
    user = repository.authenticate(username, password)
    if not user:
        return RedirectResponse("/web/login?error=Неверные+данные", status_code=303)
    token = create_session_token({"id": user["id"], "username": user["username"], "role": user["role"]})
    response = RedirectResponse("/web", status_code=303)
    response.set_cookie("app_session", token, httponly=True, samesite="lax", max_age=60 * 60 * 12)
    return response


@router.get("/web", response_class=HTMLResponse)
def dashboard(user: dict = Depends(get_current_user)) -> HTMLResponse:
    production = ProductionService().list_plans()
    managers = CommercialService().list_managers()
    body = (
        _nav(user)
        + "<h1>Внутренний кабинет</h1>"
        + f'<div class="card"><strong>Менеджеров:</strong> {len(managers)}</div>'
        + f'<div class="card"><strong>Планов:</strong> {len(production.get("plans", []))}</div>'
        + '<div class="card"><a href="/web/managers">Открыть менеджеров</a></div>'
        + '<div class="card"><a href="/web/offers">Открыть КП</a></div>'
        + '<div class="card"><a href="/web/production">Открыть производство</a></div>'
    )
    return _page("Dashboard", body)


@router.get("/web/managers", response_class=HTMLResponse)
def managers_page(user: dict = Depends(require_roles("admin", "manager", "production"))) -> HTMLResponse:
    items = CommercialService().list_managers()
    rows = "".join(
        f"<tr><td>{item.get('id')}</td><td>{escape(item.get('fio', ''))}</td><td>{escape(item.get('contact_number', ''))}</td><td>{escape(item.get('email', ''))}</td></tr>"
        for item in items
    )
    body = _nav(user) + "<h1>Менеджеры</h1>" + f"<table><tr><th>ID</th><th>ФИО</th><th>Телефон</th><th>Email</th></tr>{rows}</table>"
    return _page("Managers", body)


@router.get("/web/offers", response_class=HTMLResponse)
def offers_page(user: dict = Depends(require_roles("admin", "manager"))) -> HTMLResponse:
    offers = ProductionService().kp_repository.list_offers(limit=100)
    rows = "".join(
        f"<tr><td>{item.get('kp_id')}</td><td>{escape(str(item.get('creation_date', '')))}</td><td>{escape(item.get('customer_name', '') or '')}</td><td>{escape(item.get('manager_name', '') or '')}</td><td>{escape(item.get('status', '') or '')}</td><td>{item.get('total_amount') or 0}</td></tr>"
        for item in offers
    )
    body = (
        _nav(user)
        + "<h1>Коммерческие предложения</h1>"
        + '<div class="card"><a href="/web/offers/new">Создать КП</a></div>'
        + f"<table><tr><th>ID</th><th>Дата</th><th>Клиент</th><th>Менеджер</th><th>Статус</th><th>Сумма</th></tr>{rows}</table>"
    )
    return _page("Offers", body)


@router.get("/web/offers/new", response_class=HTMLResponse)
def new_offer_page(user: dict = Depends(require_roles("admin", "manager"))) -> HTMLResponse:
    managers = CommercialService().list_managers()
    return _render_offer_form(user=user, managers=managers)


@router.post("/web/offers/new", response_model=None)
async def new_offer_submit(
    user: dict = Depends(require_roles("admin", "manager")),
    text: str = Form(default=""),
    manager_id: str = Form(default=""),
    client_name: str = Form(default=""),
    discount_percent: str = Form(default="0"),
    delivery_conditions: str = Form(default=""),
    payment_conditions: str = Form(default=""),
    image: UploadFile | None = File(default=None),
) -> HTMLResponse | RedirectResponse:
    managers = CommercialService().list_managers()
    values = {
        "text": text,
        "manager_id": manager_id,
        "client_name": client_name,
        "discount_percent": discount_percent,
        "delivery_conditions": delivery_conditions,
        "payment_conditions": payment_conditions,
    }

    try:
        parsed_manager_id = int(manager_id)
    except ValueError:
        return _render_offer_form(user=user, managers=managers, error="Выберите менеджера.", values=values)

    try:
        parsed_discount = float((discount_percent or "0").replace(",", "."))
    except ValueError:
        return _render_offer_form(
            user=user,
            managers=managers,
            error="Скидка должна быть числом.",
            values=values,
        )

    if image and image.content_type and not image.content_type.startswith("image/"):
        return _render_offer_form(
            user=user,
            managers=managers,
            error="Поддерживаются только изображения.",
            values=values,
        )

    workflow = CommercialWorkflowService()
    try:
        draft = await workflow.create_draft_from_form(
            text=text,
            image_bytes=await image.read() if image else None,
            image_filename=image.filename if image else None,
            manager_id=parsed_manager_id,
            client_name=client_name,
            discount_percent=parsed_discount,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
        )
    except (PlateParseError, ValueError) as exc:
        return _render_offer_form(user=user, managers=managers, error=str(exc), values=values)
    return RedirectResponse(f"/web/offers/drafts/{draft['draft_id']}", status_code=303)


@router.get("/web/offers/drafts/{draft_id}", response_class=HTMLResponse)
def offer_draft_page(
    draft_id: str,
    user: dict = Depends(require_roles("admin", "manager")),
) -> HTMLResponse:
    workflow = CommercialWorkflowService()
    try:
        draft = workflow.get_draft_details(draft_id)
    except FileNotFoundError:
        return _page("Draft not found", _nav(user) + "<h1>Черновик не найден</h1>")
    return _render_offer_preview(user=user, draft=draft)


@router.post("/web/offers/drafts/{draft_id}/generate-files")
def generate_offer_draft_files(
    draft_id: str,
    user: dict = Depends(require_roles("admin", "manager")),
) -> RedirectResponse:
    workflow = CommercialWorkflowService()
    try:
        workflow.generate_files(draft_id)
    except (FileNotFoundError, ValueError):
        return RedirectResponse("/web/offers", status_code=303)
    return RedirectResponse(f"/web/offers/drafts/{draft_id}", status_code=303)


@router.post("/web/offers/drafts/{draft_id}/save")
def save_offer_draft(
    draft_id: str,
    user: dict = Depends(require_roles("admin", "manager")),
) -> RedirectResponse:
    workflow = CommercialWorkflowService()
    try:
        workflow.save_offer(draft_id)
    except FileNotFoundError:
        return RedirectResponse("/web/offers", status_code=303)
    return RedirectResponse(f"/web/offers/drafts/{draft_id}", status_code=303)


@router.get("/web/production", response_class=HTMLResponse)
def production_page(user: dict = Depends(require_roles("admin", "production"))) -> HTMLResponse:
    plans = ProductionService().list_plans()
    rows = "".join(
        f"<tr><td>{escape(plan.get('id', ''))}</td><td>{escape(plan.get('name', ''))}</td><td>{escape(plan.get('start_date', ''))}</td><td>{plan.get('total_days', 0)}</td><td>{plan.get('total_tracks', 0)}</td></tr>"
        for plan in plans.get("plans", [])
    )
    body = _nav(user) + "<h1>Планы производства</h1>" + f"<table><tr><th>ID</th><th>Название</th><th>Старт</th><th>Дней</th><th>Дорожек</th></tr>{rows}</table>"
    return _page("Production", body)

