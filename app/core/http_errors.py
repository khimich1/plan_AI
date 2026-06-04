"""Safe HTTP error responses: log server-side, generic client-facing messages."""

from __future__ import annotations

import logging
from typing import NoReturn

from fastapi import HTTPException, status

_log = logging.getLogger("app.api.commercial")

MSG_PARSE_FAILED = "Не удалось обработать ввод. Проверьте формат данных."
MSG_VALIDATION = "Проверьте введённые данные."
MSG_INTERNAL = "Внутренняя ошибка сервера. Повторите попытку позже."


def raise_parse_client_error(exc: BaseException, *, where: str) -> NoReturn:
    _log.warning("%s: parse error", where, exc_info=exc)
    raise HTTPException(
        status_code=status.HTTP_400_BAD_REQUEST,
        detail=MSG_PARSE_FAILED,
    ) from None


def raise_validation_client_error(
    exc: BaseException,
    *,
    where: str,
    detail: str | None = None,
) -> NoReturn:
    if detail is not None:
        _log.warning("%s: validation error (%s)", where, detail, exc_info=exc)
    else:
        _log.warning("%s: validation error", where, exc_info=exc)
    raise HTTPException(
        status_code=status.HTTP_400_BAD_REQUEST,
        detail=MSG_VALIDATION,
    ) from None


def raise_unexpected_server_error(_exc: BaseException, *, where: str) -> NoReturn:
    _log.exception("%s: unexpected error", where)
    raise HTTPException(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        detail=MSG_INTERNAL,
    ) from None
