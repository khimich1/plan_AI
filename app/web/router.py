from __future__ import annotations

from html import escape

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

from app.dependencies.auth import get_current_user, require_roles
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from app.services.commercial_service import CommercialService
from app.services.production_service import ProductionService

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
    body = _nav(user) + "<h1>Коммерческие предложения</h1>" + f"<table><tr><th>ID</th><th>Дата</th><th>Клиент</th><th>Менеджер</th><th>Статус</th><th>Сумма</th></tr>{rows}</table>"
    return _page("Offers", body)


@router.get("/web/production", response_class=HTMLResponse)
def production_page(user: dict = Depends(require_roles("admin", "production"))) -> HTMLResponse:
    plans = ProductionService().list_plans()
    rows = "".join(
        f"<tr><td>{escape(plan.get('id', ''))}</td><td>{escape(plan.get('name', ''))}</td><td>{escape(plan.get('start_date', ''))}</td><td>{plan.get('total_days', 0)}</td><td>{plan.get('total_tracks', 0)}</td></tr>"
        for plan in plans.get("plans", [])
    )
    body = _nav(user) + "<h1>Планы производства</h1>" + f"<table><tr><th>ID</th><th>Название</th><th>Старт</th><th>Дней</th><th>Дорожек</th></tr>{rows}</table>"
    return _page("Production", body)

