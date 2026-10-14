"""Разбор ответа 1С ``otvet_*.json``.

Живой файл содержит ``Номер_КП`` (номер КП в приложении), ``Номер_ЗК``,
``УИД_ЗК`` и ``УИД_КП``. Поля ``НомерВПриложении`` в ответе нет.
"""

from __future__ import annotations

from dataclasses import dataclass


class InvoiceReplyError(ValueError):
    """В файле нет номера КП или одного из трёх идентификаторов."""


@dataclass(frozen=True)
class InvoiceReply:
    kp_id: int
    order_number: str
    uid_order: str
    uid_kp: str


def _text(payload: dict, key: str) -> str:
    value = payload.get(key)
    if value is None:
        return ""
    return str(value).strip()


def parse_invoice_reply(payload: object) -> InvoiceReply:
    if not isinstance(payload, dict):
        raise InvoiceReplyError("ответ 1С не объект")
    raw_kp = _text(payload, "Номер_КП")
    if not raw_kp.isdigit():
        raise InvoiceReplyError("в ответе нет номера КП")
    kp_id = int(raw_kp)
    if kp_id <= 0:
        raise InvoiceReplyError("в ответе нет номера КП")
    order_number = _text(payload, "Номер_ЗК")
    uid_order = _text(payload, "УИД_ЗК")
    uid_kp = _text(payload, "УИД_КП")
    if not order_number or not uid_order or not uid_kp:
        raise InvoiceReplyError("в ответе нет номера счёта или УИД")
    return InvoiceReply(
        kp_id=kp_id,
        order_number=order_number,
        uid_order=uid_order,
        uid_kp=uid_kp,
    )
