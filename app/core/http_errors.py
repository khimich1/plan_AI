"""Safe HTTP error responses: log server-side, generic client-facing messages."""

from __future__ import annotations

import logging
from typing import NoReturn

from fastapi import HTTPException, status

from app.schemas.errors import ApiErrorBody, ERROR_CODE_UNPRICED_PLATES

_log = logging.getLogger("app.api.commercial")

MSG_PARSE_FAILED = "Не удалось обработать ввод. Проверьте формат данных."
MSG_VALIDATION = "Проверьте введённые данные."
MSG_INTERNAL = "Внутренняя ошибка сервера. Повторите попытку позже."
MSG_NOT_FOUND = "Запрошенный ресурс не найден."
MSG_ARCHIVE_NOT_FOUND = "КП не найдено."
MSG_DAY_NOT_FOUND = "Данные за указанную дату не найдены."
MSG_PLAN_VERSION_CONFLICT = "План был изменён другим запросом. Обновите страницу и повторите."
MSG_UNPROCESSABLE = "Не удалось выполнить операцию. Проверьте введённые данные."
MSG_DESTRUCTIVE_DB_BLOCKED = "Операция обнуления базы данных запрещена в текущем окружении."
MSG_TRACK_REMOVAL_FAILED = "Не удалось удалить дорожку из плана."

_TRACK_REMOVAL_CLIENT_MESSAGES: dict[str, str] = {
    "plan_not_found": "План не найден.",
    "day_not_found": MSG_DAY_NOT_FOUND,
    "day_already_completed": "День уже завершён — удаление дорожки невозможно.",
    "invalid_track_index": "Недопустимый номер дорожки.",
    "no_plate_identity": "В дорожке не найдено плит для возврата в производство.",
    "incomplete_return": "Не удалось полностью вернуть плиты в производство.",
    "db_return_failed": "Не удалось вернуть плиты в производство.",
    "plan_save_failed": "Не удалось сохранить план после удаления дорожки.",
}


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


def raise_client_error(
    exc: BaseException,
    *,
    status_code: int,
    detail: str,
    where: str,
) -> NoReturn:
    if status_code >= status.HTTP_500_INTERNAL_SERVER_ERROR:
        _log.error("%s: client error %s", where, detail, exc_info=exc)
    else:
        _log.warning("%s: client error %s", where, detail, exc_info=exc)
    raise HTTPException(status_code=status_code, detail=detail) from None


def raise_not_found_client_error(
    exc: BaseException,
    *,
    where: str,
    detail: str = MSG_NOT_FOUND,
) -> NoReturn:
    raise_client_error(
        exc,
        status_code=status.HTTP_404_NOT_FOUND,
        detail=detail,
        where=where,
    )


def raise_bad_request_client_error(
    exc: BaseException,
    *,
    where: str,
    detail: str = MSG_VALIDATION,
) -> NoReturn:
    raise_client_error(
        exc,
        status_code=status.HTTP_400_BAD_REQUEST,
        detail=detail,
        where=where,
    )


def raise_unprocessable_client_error(
    exc: BaseException,
    *,
    where: str,
    detail: str = MSG_UNPROCESSABLE,
) -> NoReturn:
    raise_client_error(
        exc,
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        detail=detail,
        where=where,
    )


def raise_structured_error(
    *,
    status_code: int,
    code: str,
    message: str,
    details: dict | None = None,
    where: str,
) -> NoReturn:
    body = ApiErrorBody(code=code, message=message, details=details)
    payload = body.model_dump(exclude_none=True)
    if status_code >= status.HTTP_500_INTERNAL_SERVER_ERROR:
        _log.error("%s: structured error %s — %s", where, code, message)
    else:
        _log.warning("%s: structured error %s — %s", where, code, message)
    raise HTTPException(status_code=status_code, detail=payload) from None


def raise_destructive_db_blocked_error(exc: BaseException, *, where: str) -> NoReturn:
    _log.warning("%s: destructive db reset blocked", where, exc_info=exc)
    raise HTTPException(
        status_code=status.HTTP_403_FORBIDDEN,
        detail=MSG_DESTRUCTIVE_DB_BLOCKED,
    ) from None


def raise_track_removal_client_error(
    exc: BaseException,
    *,
    where: str,
    status_code: int,
    code: str | None = None,
) -> NoReturn:
    detail = _TRACK_REMOVAL_CLIENT_MESSAGES.get(code or "", MSG_TRACK_REMOVAL_FAILED)
    raise_client_error(exc, status_code=status_code, detail=detail, where=where)


def raise_unpriced_plates_error(exc: BaseException, *, where: str) -> NoReturn:
    from core.exceptions import UnpricedPlatesError

    if not isinstance(exc, UnpricedPlatesError):
        raise TypeError("expected UnpricedPlatesError") from exc
    raise_structured_error(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        code=ERROR_CODE_UNPRICED_PLATES,
        message="Нет цен для части позиций",
        details={"positions": exc.positions},
        where=where,
    )
