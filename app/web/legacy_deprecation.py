from __future__ import annotations

import logging
from urllib.parse import quote

from fastapi.responses import HTMLResponse, RedirectResponse
from starlette.responses import Response

logger = logging.getLogger(__name__)

LEGACY_WEB_DEPRECATION_MSG = (
    "Legacy HTML web UI routes under /web/* are deprecated; "
    "use the React SPA under /commercial-offer/."
)

DEPRECATION_HEADER = "Deprecation"
DEPRECATION_HEADER_VALUE = "true"

SPA_BASE = "/commercial-offer"
SPA_LOGIN = f"{SPA_BASE}/login"
SPA_NEW = f"{SPA_BASE}/new"
SPA_ARCHIVE = f"{SPA_BASE}/archive"
SPA_PRODUCTION = f"{SPA_BASE}/production"

LEGACY_DRAFT_MIGRATION_NOTICE = (
    "Старый HTML-интерфейс черновиков снят. Продолжите работу в новом мастере КП."
)


def default_spa_home_for_role(role: str | None) -> str:
    if role == "production":
        return SPA_PRODUCTION
    return SPA_NEW


def spa_new_with_notice(message: str) -> str:
    return f"{SPA_NEW}?notice={quote(message, safe='')}"

def spa_new_with_error(message: str) -> str:
    return f"{SPA_NEW}?error={quote(message, safe='')}"


def spa_draft_url(draft_id: str, *, legacy: bool = True) -> str:
    query = f"draft={quote(draft_id, safe='')}"
    if legacy:
        query += "&legacy=1"
    return f"{SPA_NEW}?{query}"


def log_legacy_route_access(path: str, *, redirect_to: str | None = None) -> None:
    if redirect_to:
        logger.warning(
            "Legacy web route %s accessed; redirecting to %s. %s",
            path,
            redirect_to,
            LEGACY_WEB_DEPRECATION_MSG,
        )
    else:
        logger.warning(
            "Legacy web route %s accessed (no SPA redirect yet). %s",
            path,
            LEGACY_WEB_DEPRECATION_MSG,
        )


def apply_deprecation_headers(response: Response, *, successor: str | None = None) -> Response:
    response.headers[DEPRECATION_HEADER] = DEPRECATION_HEADER_VALUE
    if successor:
        response.headers["Link"] = f'<{successor}>; rel="successor-version"'
    return response


def deprecated_redirect(
    url: str,
    *,
    legacy_path: str,
    status_code: int = 303,
) -> RedirectResponse:
    log_legacy_route_access(legacy_path, redirect_to=url)
    response = RedirectResponse(url, status_code=status_code)
    return apply_deprecation_headers(response, successor=url)


def mark_legacy_response(response: Response, *, legacy_path: str, successor: str | None = None) -> Response:
    log_legacy_route_access(legacy_path, redirect_to=successor)
    return apply_deprecation_headers(response, successor=successor)
