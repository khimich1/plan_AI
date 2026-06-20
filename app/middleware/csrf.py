"""CSRF protection for cookie-based auth (double-submit cookie + header/form field)."""

from __future__ import annotations

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse, Response

from app.security.csrf import (
    CSRF_COOKIE_NAME,
    CSRF_FORM_FIELD,
    CSRF_HEADER_NAME,
    clear_csrf_cookie,
    generate_csrf_token,
    is_safe_method,
    set_csrf_cookie,
    tokens_match,
)


class CsrfMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next) -> Response:
        if not is_safe_method(request.method):
            cookie_token = request.cookies.get(CSRF_COOKIE_NAME)
            submitted = await self._submitted_token(request)
            if not tokens_match(cookie_token, submitted):
                return JSONResponse(
                    status_code=403,
                    content={"detail": "CSRF token missing or invalid."},
                )

        response = await call_next(request)
        self._ensure_csrf_cookie(request, response)
        return response

    async def _submitted_token(self, request: Request) -> str | None:
        header = request.headers.get(CSRF_HEADER_NAME)
        if header and header.strip():
            return header.strip()

        content_type = (request.headers.get("content-type") or "").lower()
        if "application/x-www-form-urlencoded" in content_type or "multipart/form-data" in content_type:
            form = await request.form()
            field = form.get(CSRF_FORM_FIELD)
            if field is not None:
                value = str(field).strip()
                return value or None
        return None

    def _ensure_csrf_cookie(self, request: Request, response: Response) -> None:
        if request.cookies.get(CSRF_COOKIE_NAME):
            return
        set_csrf_cookie(response, generate_csrf_token())
