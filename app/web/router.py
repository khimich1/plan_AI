from __future__ import annotations

import json
from html import escape
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile, status
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse

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
)
from app.dependencies.auth import REQUIRE_ADMIN_OR_MANAGER, get_current_user, require_roles
from app.dependencies.commercial_draft import check_draft_ownership
from app.dependencies.plate_context import get_plate_order_context
from app.repositories.auth_repository import AuthRepository
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
        <h1>Frontend build не найден</h1>
        <p>Соберите React-приложение в директорию <code>frontend/dist</code>, затем откройте страницу снова.</p>
        <p>Для локальной разработки используйте Vite dev server из директории <code>frontend/</code>.</p>
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


def _nav(user: dict) -> str:
    return (
        f"<nav>"
        f"<strong>{escape(user['username'])}</strong> ({escape(user['role'])}) "
        f'<a href="/web">Главная</a>'
        f'<a href="/web/managers">Менеджеры</a>'
        f'<a href="/web/offers">КП</a>'
        f'<a href="/web/production">Производство</a>'
        f'<a href="/web/logout">Выйти</a>'
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
        response = RedirectResponse(f"{SPA_LOGIN}?error=Неверные+данные", status_code=303)
        return mark_legacy_response(response, legacy_path="/web/login", successor=SPA_LOGIN)
    token = create_session_token({"id": user["id"], "username": user["username"], "role": user["role"]})
    home = default_spa_home_for_role(user["role"])
    response = RedirectResponse(home, status_code=303)
    set_session_cookie(response, token)
    return mark_legacy_response(response, legacy_path="/web/login", successor=home)


@router.get("/web/logout")
def web_logout() -> RedirectResponse:
    response = RedirectResponse(SPA_LOGIN, status_code=303)
    clear_session_cookie(response)
    return mark_legacy_response(response, legacy_path="/web/logout", successor=SPA_LOGIN)


@router.get("/web")
def dashboard(user: dict = Depends(get_current_user)) -> RedirectResponse:
    target = default_spa_home_for_role(user.get("role"))
    return deprecated_redirect(target, legacy_path="/web")


def _legacy_dashboard_html(user: dict) -> HTMLResponse:
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


@router.get("/web/managers")
def managers_page(
    user: dict = Depends(require_roles("admin", "manager", "production")),
) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_ARCHIVE, legacy_path="/web/managers")


def _legacy_managers_page_html(user: dict) -> HTMLResponse:
    role = user.get("role", "")
    role_json = json.dumps(role)
    create_card_hidden = "" if role in {"admin", "manager"} else ' style="display:none"'
    body = (
        _nav(user)
        + f"""
    <h1>Менеджеры и КП</h1>
    <div class="card" id="create-offer-card"{create_card_hidden}>
      <h2>Создать КП</h2>
      <p style="margin-top: 0;">Шаги: список плит -> менеджер -> клиент -> скидка/условия -> превью -> сохранить.</p>
      <label>Список плит:</label><br>
      <textarea id="plates-text" rows="7" style="width:100%;"></textarea><br><br>

      <div style="display:grid; grid-template-columns: 1fr 1fr; gap: 12px;">
        <div>
          <label>Менеджер:</label><br>
          <select id="manager-id" style="width:100%;"></select>
        </div>
        <div>
          <label>Клиент:</label><br>
          <input id="customer-name" type="text" style="width:100%;">
        </div>
        <div>
          <label>Скидка (%):</label><br>
          <input id="discount-percent" type="number" min="0" max="100" step="0.1" value="0" style="width:100%;">
        </div>
        <div>
          <label>Срок для «в работе»:</label><br>
          <input id="execution-terms-input" type="text" placeholder="например: 14 дней / 2 недели / 01.05.2026" style="width:100%;">
        </div>
      </div>
      <br>
      <label>Условия поставки:</label><br>
      <input id="delivery-conditions" type="text" style="width:100%;"><br><br>
      <label>Условия оплаты:</label><br>
      <input id="payment-conditions" type="text" style="width:100%;"><br><br>

      <button id="btn-preview">Сгенерировать превью</button>
      <button id="btn-recognize-screen">Распознать СКРИН</button>
      <input id="screen-file-input" type="file" accept="image/*" style="display:none;">
      <button id="btn-preview-check-xlsx" disabled>XLSX проверка</button>
      <button id="btn-preview-xlsx" disabled>Скачать превью XLSX</button>
      <button id="btn-save-work" disabled>Сохранить в БД (в работе)</button>
      <button id="btn-save-archive" disabled>В архив</button>
      <p id="create-status"></p>
      <pre id="recognized-output" style="background:#f4f7ff; padding:10px; white-space:pre-wrap;"></pre>
      <pre id="preview-output" style="background:#f8f8f8; padding:10px; white-space:pre-wrap;"></pre>
    </div>

    <div class="card">
      <h2>Архив КП</h2>
      <div style="display:flex; gap:8px; flex-wrap: wrap; margin-bottom: 8px;">
        <button data-status="archived" class="status-btn">В архиве</button>
        <button data-status="in_production" class="status-btn">В производстве</button>
        <button data-status="completed" class="status-btn">Выполненные</button>
        <button data-status="all" class="status-btn">Все</button>
      </div>
      <div style="display:flex; gap:8px; margin-bottom: 8px;">
        <input id="search-kp-id" type="number" min="1" placeholder="Номер КП" style="max-width:160px;">
        <button id="btn-search-kp">Найти по номеру</button>
      </div>
      <p id="archive-status"></p>
      <table id="offers-table">
        <thead>
          <tr>
            <th>ID</th><th>Дата</th><th>Клиент</th><th>Менеджер</th><th>Статус</th><th>Готовность</th><th>Сумма</th><th></th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>

    <div class="card" id="offer-details-card">
      <h2>Карточка КП</h2>
      <p id="details-empty">Выберите КП из таблицы выше.</p>
      <div id="offer-details"></div>
    </div>
    <script>
    const USER_ROLE = """
        + role_json
        + """;
    const canEdit = USER_ROLE === "admin" || USER_ROLE === "manager";
    const state = {
      managers: [],
      managersById: {},
      preview: null,
      recognizedText: "",
      activeStatus: "archived",
      selectedOfferId: null
    };

    const createStatusEl = document.getElementById("create-status");
    const previewOutputEl = document.getElementById("preview-output");
    const managerSelectEl = document.getElementById("manager-id");
    const btnPreviewEl = document.getElementById("btn-preview");
    const btnRecognizeScreenEl = document.getElementById("btn-recognize-screen");
    const screenFileInputEl = document.getElementById("screen-file-input");
    const btnPreviewCheckXlsxEl = document.getElementById("btn-preview-check-xlsx");
    const btnPreviewXlsxEl = document.getElementById("btn-preview-xlsx");
    const btnSaveWorkEl = document.getElementById("btn-save-work");
    const btnSaveArchiveEl = document.getElementById("btn-save-archive");
    const archiveStatusEl = document.getElementById("archive-status");
    const offersBodyEl = document.querySelector("#offers-table tbody");
    const offerDetailsEl = document.getElementById("offer-details");
    const detailsEmptyEl = document.getElementById("details-empty");
    const createCardEl = document.getElementById("create-offer-card");
    const recognizedOutputEl = document.getElementById("recognized-output");

    function toDateRu(value) {
      const d = value ? new Date(value) : new Date();
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = String(d.getFullYear());
      return `${dd}.${mm}.${yyyy}`;
    }

    function esc(value) {
      return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
    }

    async function api(path, options = {}) {
      const cfg = {
        credentials: "include",
        ...options,
        headers: {
          "Content-Type": "application/json",
          ...(options.headers || {})
        }
      };
      if (!cfg.body) {
        delete cfg.headers["Content-Type"];
      }
      const res = await fetch(path, cfg);
      let data = null;
      try {
        data = await res.json();
      } catch (_err) {
        data = null;
      }
      if (!res.ok) {
        const detail = (data && (data.detail || data.message)) || `HTTP ${res.status}`;
        throw new Error(detail);
      }
      return data;
    }

    function setCreateStatus(text, isError = false) {
      createStatusEl.textContent = text;
      createStatusEl.style.color = isError ? "red" : "inherit";
    }

    function setArchiveStatus(text, isError = false) {
      archiveStatusEl.textContent = text;
      archiveStatusEl.style.color = isError ? "red" : "inherit";
    }

    function renderManagers() {
      if (!state.managers.length) {
        managerSelectEl.innerHTML = "<option value=''>Нет менеджеров</option>";
        return;
      }
      managerSelectEl.innerHTML = state.managers
        .map((m) => `<option value="${m.id}">${esc(m.fio || "Без имени")}</option>`)
        .join("");
    }

    async function loadManagers() {
      const data = await api("/api/v1/managers");
      state.managers = data.items || [];
      state.managersById = {};
      for (const item of state.managers) {
        state.managersById[String(item.id)] = item;
      }
      renderManagers();
    }

    function currentCreatePayload() {
      const managerId = managerSelectEl.value;
      const manager = state.managersById[String(managerId)];
      if (!manager) {
        throw new Error("Выберите менеджера");
      }
      const customerName = document.getElementById("customer-name").value.trim();
      if (!customerName) {
        throw new Error("Введите клиента");
      }
      const discountPercent = Number(document.getElementById("discount-percent").value || 0);
      if (Number.isNaN(discountPercent) || discountPercent < 0 || discountPercent > 100) {
        throw new Error("Скидка должна быть от 0 до 100");
      }
      const platesText = document.getElementById("plates-text").value.trim();
      if (!platesText) {
        throw new Error("Введите список плит");
      }
      return {
        manager,
        customerName,
        discountPercent,
        platesText,
        deliveryConditions: document.getElementById("delivery-conditions").value.trim(),
        paymentConditions: document.getElementById("payment-conditions").value.trim(),
        executionTermsInput: document.getElementById("execution-terms-input").value.trim()
      };
    }

    function currentPlatesText() {
      return document.getElementById("plates-text").value.trim();
    }

    function renderPreview(data) {
      const warnings = data.warnings || [];
      const unparsed = data.unparsed_lines || [];
      const optimization = data.optimization || {};
      previewOutputEl.textContent =
        "Превью готово\\n" +
        `Плит в расчете: ${optimization.total_plates ?? "-"}\\n` +
        `Сумма: ${Number(data.total_sum || 0).toLocaleString("ru-RU")} ₽\\n` +
        `Неразобранные строки: ${unparsed.length}\\n` +
        `Предупреждения: ${warnings.length}\\n\\n` +
        (warnings.length ? `Warnings:\\n- ${warnings.join("\\n- ")}\\n\\n` : "") +
        (unparsed.length ? `Unparsed:\\n- ${unparsed.join("\\n- ")}` : "");
    }

    async function createPreview() {
      try {
        setCreateStatus("Генерирую превью...");
        const payload = currentCreatePayload();
        const data = await api("/api/v1/commercial/generate-preview", {
          method: "POST",
          body: JSON.stringify({ text: payload.platesText })
        });
        state.preview = data;
        renderPreview(data);
        btnPreviewXlsxEl.disabled = !canEdit;
        btnSaveWorkEl.disabled = !canEdit;
        btnSaveArchiveEl.disabled = !canEdit;
        setCreateStatus("Превью успешно сгенерировано.");
      } catch (err) {
        btnPreviewXlsxEl.disabled = true;
        setCreateStatus(err.message, true);
      }
    }

    function renderRecognizedText(data) {
      const lines = data.lines || [];
      const warnings = data.warnings || [];
      recognizedOutputEl.textContent =
        "Результат OCR (отдельно, без автоподстановки):\\n" +
        `Метод: ${data.method || "-"}\\n` +
        `Уверенность: ${Number(data.confidence || 0).toFixed(2)}\\n` +
        `Строк: ${lines.length}\\n\\n` +
        (lines.length ? lines.join("\\n") + "\\n\\n" : "") +
        (warnings.length ? `Warnings:\\n- ${warnings.join("\\n- ")}` : "");
    }

    async function recognizeScreenFile(file) {
      if (!file) return;
      try {
        setCreateStatus("Распознаю скрин...");
        const formData = new FormData();
        formData.append("image", file);
        const res = await fetch("/api/v1/commercial/recognize-screen", {
          method: "POST",
          credentials: "include",
          body: formData
        });
        let data = null;
        try {
          data = await res.json();
        } catch (_err) {
          data = null;
        }
        if (!res.ok) {
          const detail = (data && (data.detail || data.message)) || `HTTP ${res.status}`;
          throw new Error(detail);
        }
        state.recognizedText = (data.recognized_text || "").trim();
        renderRecognizedText(data);
        btnPreviewCheckXlsxEl.disabled = !(currentPlatesText() && state.recognizedText);
        setCreateStatus("Скрин распознан. Текст показан отдельно ниже.");
      } catch (err) {
        setCreateStatus(err.message, true);
      } finally {
        screenFileInputEl.value = "";
      }
    }

    async function downloadPreviewCheckXlsx() {
      try {
        const platesText = currentPlatesText();
        if (!platesText) {
          throw new Error("Сначала заполните поле «Список плит».");
        }
        if (!state.recognizedText) {
          throw new Error("Сначала нажмите «Распознать СКРИН».");
        }
        setCreateStatus("Формирую XLSX проверку...");
        const res = await fetch("/api/v1/commercial/preview-check-xlsx", {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            plates_text: platesText,
            recognized_text: state.recognizedText
          })
        });
        if (!res.ok) {
          let detail = `HTTP ${res.status}`;
          try {
            const data = await res.json();
            detail = data.detail || detail;
          } catch (_err) {}
          throw new Error(detail);
        }
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `preview_check_${Date.now()}.xlsx`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        setCreateStatus("XLSX проверка скачан.");
      } catch (err) {
        setCreateStatus(err.message, true);
      }
    }

    async function downloadPreviewXlsx() {
      try {
        if (!state.preview || !state.preview.draft_id) {
          throw new Error("Сначала сгенерируйте превью.");
        }
        const payload = currentCreatePayload();
        setCreateStatus("Готовлю XLSX превью...");
        const res = await fetch(`/api/v1/commercial/drafts/${state.preview.draft_id}/xlsx`, {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            customer_name: payload.customerName,
            manager_name: payload.manager.fio || "",
            manager_phone: payload.manager.contact_number || "",
            manager_email: payload.manager.email || "",
            discount_percent: payload.discountPercent,
            delivery_conditions: payload.deliveryConditions,
            payment_conditions: payload.paymentConditions
          })
        });
        if (!res.ok) {
          let detail = `HTTP ${res.status}`;
          try {
            const data = await res.json();
            detail = data.detail || detail;
          } catch (_err) {}
          throw new Error(detail);
        }
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `KP_preview_${state.preview.draft_id.slice(0, 8)}.xlsx`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        setCreateStatus("XLSX превью скачан.");
      } catch (err) {
        setCreateStatus(err.message, true);
      }
    }

    async function saveOffer(mode) {
      try {
        if (!canEdit) {
          throw new Error("Недостаточно прав для сохранения.");
        }
        if (!state.preview || !state.preview.order_data || !state.preview.order_data.length) {
          throw new Error("Сначала сгенерируйте превью.");
        }
        setCreateStatus("Сохраняю КП...");
        const payload = currentCreatePayload();
        const request = {
          creation_date: toDateRu(new Date()),
          customer_name: payload.customerName,
          manager_name: payload.manager.fio || "",
          manager_phone: payload.manager.contact_number || "",
          manager_email: payload.manager.email || "",
          discount_percent: payload.discountPercent,
          delivery_conditions: payload.deliveryConditions,
          payment_conditions: payload.paymentConditions,
          execution_terms_input: payload.executionTermsInput,
          save_mode: mode,
          order_data: state.preview.order_data
        };
        const result = await api("/api/v1/offers", {
          method: "POST",
          body: JSON.stringify(request)
        });
        setCreateStatus(`КП сохранено: №${result.kp_id} (${result.status}).`);
        await loadOffers(mode === "archive" ? "archived" : "in_production");
      } catch (err) {
        setCreateStatus(err.message, true);
      }
    }

    function statusText(status) {
      if (status === "в архиве") return "📦 в архиве";
      if (status === "в работе") return "🏭 в работе";
      if (status === "выполнено") return "✅ выполнено";
      return status || "—";
    }

    function renderOffers(items) {
      if (!items.length) {
        offersBodyEl.innerHTML = "<tr><td colspan='8'>Нет данных</td></tr>";
        return;
      }
      offersBodyEl.innerHTML = items
        .map((item) => `
          <tr>
            <td>${item.kp_id}</td>
            <td>${esc(item.creation_date || "")}</td>
            <td>${esc(item.customer_name || "")}</td>
            <td>${esc(item.manager_name || "")}</td>
            <td>${statusText(item.status)}</td>
            <td>${Number(item.completion_percentage || 0).toFixed(1)}%</td>
            <td>${Number(item.total_amount || 0).toLocaleString("ru-RU")} ₽</td>
            <td><button data-open-offer="${item.kp_id}">Открыть</button></td>
          </tr>
        `)
        .join("");
    }

    async function loadOffers(status = state.activeStatus) {
      try {
        state.activeStatus = status;
        setArchiveStatus("Загружаю список КП...");
        const data = await api(`/api/v1/offers?status=${encodeURIComponent(status)}&limit=500`);
        renderOffers(data.items || []);
        setArchiveStatus(`Загружено: ${data.count || 0}`);
      } catch (err) {
        setArchiveStatus(err.message, true);
      }
    }

    function renderDetails(item) {
      detailsEmptyEl.style.display = "none";
      const plates = item.plates || [];
      const canMove = canEdit && item.status === "в архиве";
      const canDelete = canEdit;
      const canDiscount = canEdit;
      offerDetailsEl.innerHTML = `
        <p><strong>КП №${item.kp_id}</strong></p>
        <p>Клиент: ${esc(item.customer_name || "—")}<br>
        Менеджер: ${esc(item.manager_name || "—")}<br>
        Дата: ${esc(item.creation_date || "—")}<br>
        Статус: ${statusText(item.status)}<br>
        Срок: ${esc(item.execution_terms || "—")}<br>
        Сумма: ${Number(item.total_amount || 0).toLocaleString("ru-RU")} ₽</p>
        <p><strong>Позиции (${plates.length}):</strong></p>
        <div style="max-height: 220px; overflow: auto; border: 1px solid #ddd; padding: 8px;">
          ${plates.slice(0, 20).map((p, idx) =>
            `${idx + 1}) ${esc(p.plate_name || "")} - ${Number(p.qty || 0)} шт`
          ).join("<br>")}
          ${plates.length > 20 ? "<br>... и еще позиции" : ""}
        </div>
        <br>
        <div style="display:flex; gap:8px; flex-wrap: wrap;">
          <button data-download-pdf="${item.kp_id}">PDF</button>
          <button data-download-xlsx="${item.kp_id}">XLSX</button>
          <button data-update-discount="${item.kp_id}" ${canDiscount ? "" : "disabled"}>Изменить скидку</button>
          <button data-move-production="${item.kp_id}" ${canMove ? "" : "disabled"}>В производство</button>
          <button data-delete-offer="${item.kp_id}" ${canDelete ? "" : "disabled"}>Удалить</button>
        </div>
      `;
    }

    async function openOffer(kpId) {
      try {
        setArchiveStatus(`Открываю КП №${kpId}...`);
        const item = await api(`/api/v1/offers/${kpId}`);
        state.selectedOfferId = kpId;
        renderDetails(item);
        setArchiveStatus(`Открыто КП №${kpId}`);
      } catch (err) {
        setArchiveStatus(err.message, true);
      }
    }

    async function searchOffer() {
      const input = document.getElementById("search-kp-id").value.trim();
      if (!input) {
        loadOffers(state.activeStatus);
        return;
      }
      try {
        setArchiveStatus("Ищу КП...");
        const data = await api(`/api/v1/offers?kp_id=${encodeURIComponent(input)}`);
        renderOffers(data.items || []);
        setArchiveStatus(`Найдено: ${data.count || 0}`);
      } catch (err) {
        setArchiveStatus(err.message, true);
      }
    }

    async function updateDiscount(kpId) {
      const raw = prompt("Введите скидку 0-100", "0");
      if (raw === null) return;
      try {
        const value = Number(String(raw).replace(",", "."));
        if (Number.isNaN(value)) throw new Error("Неверный формат скидки.");
        await api(`/api/v1/offers/${kpId}/discount`, {
          method: "PATCH",
          body: JSON.stringify({ discount_percent: value })
        });
        await openOffer(kpId);
        await loadOffers(state.activeStatus);
      } catch (err) {
        setArchiveStatus(err.message, true);
      }
    }

    async function moveToProduction(kpId) {
      const raw = prompt("Введите срок (например: 14 дней / 2 недели / 01.05.2026)", "14 дней");
      if (raw === null) return;
      try {
        await api(`/api/v1/offers/${kpId}/move-to-production`, {
          method: "PATCH",
          body: JSON.stringify({ execution_terms_input: raw })
        });
        await openOffer(kpId);
        await loadOffers("in_production");
      } catch (err) {
        setArchiveStatus(err.message, true);
      }
    }

    async function deleteOffer(kpId) {
      const ok = confirm(`Удалить КП №${kpId}? Это действие необратимо.`);
      if (!ok) return;
      try {
        await api(`/api/v1/offers/${kpId}`, { method: "DELETE" });
        offerDetailsEl.innerHTML = "";
        detailsEmptyEl.style.display = "block";
        await loadOffers(state.activeStatus);
      } catch (err) {
        setArchiveStatus(err.message, true);
      }
    }

    btnPreviewEl.addEventListener("click", createPreview);
    btnRecognizeScreenEl.addEventListener("click", () => screenFileInputEl.click());
    screenFileInputEl.addEventListener("change", () => {
      const file = screenFileInputEl.files && screenFileInputEl.files[0];
      recognizeScreenFile(file);
    });
    btnPreviewCheckXlsxEl.addEventListener("click", downloadPreviewCheckXlsx);
    btnPreviewXlsxEl.addEventListener("click", downloadPreviewXlsx);
    btnSaveWorkEl.addEventListener("click", () => saveOffer("work"));
    btnSaveArchiveEl.addEventListener("click", () => saveOffer("archive"));
    document.getElementById("plates-text").addEventListener("input", () => {
      btnPreviewCheckXlsxEl.disabled = !(currentPlatesText() && state.recognizedText);
    });
    document.getElementById("btn-search-kp").addEventListener("click", searchOffer);

    document.querySelectorAll(".status-btn").forEach((btn) => {
      btn.addEventListener("click", () => loadOffers(btn.dataset.status || "all"));
    });

    offersBodyEl.addEventListener("click", (event) => {
      const btn = event.target.closest("button[data-open-offer]");
      if (!btn) return;
      openOffer(btn.dataset.openOffer);
    });

    document.getElementById("offer-details-card").addEventListener("click", (event) => {
      const el = event.target;
      if (!(el instanceof HTMLElement)) return;
      const pdfId = el.getAttribute("data-download-pdf");
      const xlsxId = el.getAttribute("data-download-xlsx");
      const discountId = el.getAttribute("data-update-discount");
      const moveId = el.getAttribute("data-move-production");
      const deleteId = el.getAttribute("data-delete-offer");
      if (pdfId) {
        window.open(`/api/v1/offers/${pdfId}/pdf`, "_blank");
      } else if (xlsxId) {
        window.open(`/api/v1/offers/${xlsxId}/xlsx`, "_blank");
      } else if (discountId) {
        updateDiscount(discountId);
      } else if (moveId) {
        moveToProduction(moveId);
      } else if (deleteId) {
        deleteOffer(deleteId);
      }
    });

    async function bootstrap() {
      if (!canEdit) {
        createCardEl.style.display = "none";
      }
      state.recognizedText = "";
      recognizedOutputEl.textContent = "Распознанный текст из скрина появится здесь.";
      btnPreviewCheckXlsxEl.disabled = true;
      btnPreviewXlsxEl.disabled = true;
      btnSaveWorkEl.disabled = true;
      btnSaveArchiveEl.disabled = true;
      await loadManagers();
      await loadOffers("archived");
    }

    bootstrap();
    </script>
    """
    )
    return _page("Managers", body)


@router.get("/web/offers")
def offers_page(
    user: dict = Depends(require_roles("admin", "manager", "production")),
) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_ARCHIVE, legacy_path="/web/offers")


def _legacy_offers_page_html(user: dict) -> HTMLResponse:
    # Object-level RBAC: same filters as REST /api/v1/offers (OffersService).
    offers = OffersService().list_offers(user=user, status="all", limit=100)
    rows = "".join(
        f"<tr><td>{item.get('kp_id')}</td><td>{escape(str(item.get('creation_date', '')))}</td><td>{escape(item.get('customer_name', '') or '')}</td><td>{escape(item.get('manager_name', '') or '')}</td><td>{escape(item.get('status', '') or '')}</td><td>{item.get('total_amount') or 0}</td></tr>"
        for item in offers
    )
    body = (
        _nav(user)
        + "<h1>Коммерческие предложения</h1>"
        + '<div class="card"><a href="/commercial-offer/new">Создать КП</a></div>'
        + f"<table><tr><th>ID</th><th>Дата</th><th>Клиент</th><th>Менеджер</th><th>Статус</th><th>Сумма</th></tr>{rows}</table>"
    )
    return _page("Offers", body)


@router.get("/web/offers/new")
def new_offer_page(user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER)) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_NEW, legacy_path="/web/offers/new")


@router.post("/web/offers/new", response_model=None)
async def new_offer_submit(
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
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
        return mark_legacy_response(
            _render_offer_form(user=user, managers=managers, error="Выберите менеджера.", values=values),
            legacy_path="/web/offers/new",
            successor=SPA_NEW,
        )

    try:
        parsed_discount = float((discount_percent or "0").replace(",", "."))
    except ValueError:
        return mark_legacy_response(
            _render_offer_form(
                user=user,
                managers=managers,
                error="Скидка должна быть числом.",
                values=values,
            ),
            legacy_path="/web/offers/new",
            successor=SPA_NEW,
        )

    try:
        image_bytes, image_name = await prepare_commercial_ocr_upload(
            image=image,
            user_id=int(user["id"]),
        )
    except HTTPException as exc:
        msg = exc.detail if isinstance(exc.detail, str) else "Ошибка загрузки файла."
        return mark_legacy_response(
            _render_offer_form(user=user, managers=managers, error=msg, values=values),
            legacy_path="/web/offers/new",
            successor=SPA_NEW,
        )

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
        return mark_legacy_response(
            _render_offer_form(user=user, managers=managers, error=str(exc), values=values),
            legacy_path="/web/offers/new",
            successor=SPA_NEW,
        )
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
            spa_new_with_notice("Черновик не найден или недоступен."),
            legacy_path=legacy_path,
        )
    workflow = CommercialWorkflowService()
    try:
        workflow.get_draft_details(draft_id)
    except FileNotFoundError:
        return deprecated_redirect(
            spa_new_with_notice("Черновик не найден или недоступен."),
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
            spa_new_with_notice("Не удалось сгенерировать файлы для черновика."),
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
            spa_new_with_notice("Не удалось сохранить черновик в базу."),
            status_code=303,
        )
        return mark_legacy_response(response, legacy_path=legacy_path, successor=successor)
    response = RedirectResponse(successor, status_code=303)
    return mark_legacy_response(response, legacy_path=legacy_path, successor=successor)


@router.get("/web/production")
def production_page(user: dict = Depends(require_roles("admin", "production"))) -> RedirectResponse:
    _ = user
    return deprecated_redirect(SPA_PRODUCTION, legacy_path="/web/production")


def _legacy_production_page_html(user: dict) -> HTMLResponse:
    plans = ProductionService().list_plans()
    rows = "".join(
        f"<tr><td>{escape(plan.get('id', ''))}</td><td>{escape(plan.get('name', ''))}</td><td>{escape(plan.get('start_date', ''))}</td><td>{plan.get('total_days', 0)}</td><td>{plan.get('total_tracks', 0)}</td></tr>"
        for plan in plans.get("plans", [])
    )
    body = _nav(user) + "<h1>Планы производства</h1>" + f"<table><tr><th>ID</th><th>Название</th><th>Старт</th><th>Дней</th><th>Дорожек</th></tr>{rows}</table>"
    return _page("Production", body)


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

