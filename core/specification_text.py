"""Абзацы спецификации и пропись. Клиент эти фразы не собирает."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Literal, Mapping

PaymentKind = Literal["prepay_100", "split_50_50", "split_share", "deferral", "custom"]
TermKind = Literal["by_date", "pile_rhythm", "after_payment"]
DeliveryKind = Literal["pickup", "site"]
PileUnit = Literal["week", "day"]

SUPPLIER_NAME = "ООО «ЖБК СТАРТ»"
SUPPLIER_DIRECTOR_GENITIVE = "директора Шишова Александра Васильевича"
SUPPLIER_DIRECTOR_SHORT = "Шишов А. В."
SUPPLIER_INN = "4401082520"
SUPPLIER_KPP = "440101001"
SUPPLIER_CITY = "г. Кострома"
PICKUP_ADDRESS = (
    "склад готовой продукции, обособленное подразделение «ЖБК СТАРТ» "
    "ООО «ЖБК СТАРТ», 150020, г. Ярославль, пр. Домостроителей, дом 1, стр. 3"
)

PREPAY_CLAUSE_FULL = (
    "при условии поступления предварительной оплаты на расчетный счет поставщика "
    "в полном объеме в установленный в настоящей спецификации срок"
)
PREPAY_CLAUSE_PARTIAL = (
    "при условии поступления авансов, указанных в условиях оплаты настоящей спецификации, "
    "на расчетный счет поставщика в установленный в них срок"
)

_MONTHS = (
    "",
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
)
_ONES_M = (
    "",
    "один",
    "два",
    "три",
    "четыре",
    "пять",
    "шесть",
    "семь",
    "восемь",
    "девять",
)
_ONES_F = (
    "",
    "одна",
    "две",
    "три",
    "четыре",
    "пять",
    "шесть",
    "семь",
    "восемь",
    "девять",
)
_TEENS = (
    "десять",
    "одиннадцать",
    "двенадцать",
    "тринадцать",
    "четырнадцать",
    "пятнадцать",
    "шестнадцать",
    "семнадцать",
    "восемнадцать",
    "девятнадцать",
)
_TENS = (
    "",
    "",
    "двадцать",
    "тридцать",
    "сорок",
    "пятьдесят",
    "шестьдесят",
    "семьдесят",
    "восемьдесят",
    "девяносто",
)
_HUNDREDS = (
    "",
    "сто",
    "двести",
    "триста",
    "четыреста",
    "пятьсот",
    "шестьсот",
    "семьсот",
    "восемьсот",
    "девятьсот",
)
_SHARE_PERCENTS = (10, 20, 30, 40)
_KOPECK = Decimal("1")


class SpecificationTextError(ValueError):
    """Выбор нельзя собрать в абзац."""


@dataclass(frozen=True)
class SpecificationChoice:
    payment: PaymentKind
    term: TermKind
    delivery: DeliveryKind
    payment_date: date | None = None
    payment_days: int | None = None
    second_share_percent: int | None = None
    second_payment_date: date | None = None
    custom_text: str | None = None
    term_date: date | None = None
    pile_count: int | None = None
    pile_unit: PileUnit | None = None
    term_days: int | None = None
    delivery_address: str | None = None


@dataclass(frozen=True)
class SpecificationParagraphs:
    payment: str
    term: str
    delivery: str


def render_specification(
    choice: SpecificationChoice,
    *,
    payable_total: int,
    has_piles: bool = False,
) -> SpecificationParagraphs:
    """Оплата, срок и вид поставки от итога в копейках."""
    if isinstance(payable_total, bool) or not isinstance(payable_total, int):
        raise SpecificationTextError("Итог указывается в копейках")
    if payable_total < 0:
        raise SpecificationTextError("Сумма не может быть отрицательной")
    _validate(choice, has_piles=has_piles)
    return SpecificationParagraphs(
        payment=_payment_paragraph(choice, payable_total),
        term=_term_paragraph(choice, has_piles=has_piles),
        delivery=_delivery_paragraph(choice),
    )


def payment_tranche_kopecks(
    payment: str,
    payable_total: int,
    *,
    share_percent: int | None = None,
) -> tuple[int, ...]:
    """Доли от итога. Последний транш — остаток, сумма равна итогу."""
    total = int(payable_total)
    if payment == "split_50_50":
        first = _percent_kopecks(total, 50)
        return (first, total - first)
    if payment == "split_share":
        share = int(share_percent or 0)
        first = _percent_kopecks(total, 50)
        second = _percent_kopecks(total, share)
        if first + second > total:
            second = total - first
        return (first, second, total - first - second)
    if payment == "prepay_100":
        return (total,)
    return ()


def rubles_in_words(kopecks: int) -> str:
    """Пропись транша. Без копеек заканчивается на рубль/рубля/рублей."""
    if kopecks < 0:
        raise SpecificationTextError("Сумма не может быть отрицательной")
    rub = kopecks // 100
    kop = kopecks % 100
    text = f"{integer_in_words(rub)} {_plural(rub, 'рубль', 'рубля', 'рублей')}"
    if kop:
        text = f"{text} {kop:02d} {_plural(kop, 'копейка', 'копейки', 'копеек')}"
    return text


def total_amount_words(kopecks: int) -> str:
    """Пропись итога спецификации: с заглавной и копейками, включая 00."""
    if kopecks < 0:
        raise SpecificationTextError("Сумма не может быть отрицательной")
    rub = kopecks // 100
    kop = kopecks % 100
    text = (
        f"{integer_in_words(rub)} {_plural(rub, 'рубль', 'рубля', 'рублей')} "
        f"{kop:02d} {_plural(kop, 'копейка', 'копейки', 'копеек')}"
    )
    return text[:1].upper() + text[1:]


def integer_in_words(value: int, *, feminine: bool = False) -> str:
    if value < 0:
        raise SpecificationTextError("Число не может быть отрицательным")
    if value == 0:
        return "ноль"
    billions = value // 1_000_000_000
    millions = (value // 1_000_000) % 1000
    thousands = (value // 1000) % 1000
    rest = value % 1000
    parts: list[str] = []
    if billions:
        parts.append(_triplet(billions, feminine=False))
        parts.append(_plural(billions, "миллиард", "миллиарда", "миллиардов"))
    if millions:
        parts.append(_triplet(millions, feminine=False))
        parts.append(_plural(millions, "миллион", "миллиона", "миллионов"))
    if thousands:
        parts.append(_triplet(thousands, feminine=True))
        parts.append(_plural(thousands, "тысяча", "тысячи", "тысяч"))
    if rest or not parts:
        parts.append(_triplet(rest, feminine=feminine))
    return " ".join(part for part in parts if part)


def calendar_days_phrase(days: int) -> str:
    unit = day_unit(days)
    adjective = "календарный" if unit == "день" else "календарных"
    return f"{days} ({integer_in_words(days)}) {adjective} {unit}"


def day_unit(days: int) -> str:
    return _plural(days, "день", "дня", "дней")


def format_rub(kopecks: int) -> str:
    sign = "-" if kopecks < 0 else ""
    amount = abs(kopecks)
    rub = amount // 100
    kop = amount % 100
    grouped = f"{rub:,}".replace(",", " ")
    return f"{sign}{grouped},{kop:02d}"


def date_in_words(value: date) -> str:
    return f"{value.day} {_MONTHS[value.month]} {value.year} года"


def format_date(value: date) -> str:
    return value.strftime("%d.%m.%Y")


def money_phrase(kopecks: int) -> str:
    return f"{format_rub(kopecks)} ({rubles_in_words(kopecks)})"


def _payment_paragraph(choice: SpecificationChoice, total: int) -> str:
    if choice.payment == "prepay_100":
        return f"{money_phrase(total)} — предварительная оплата в размере 100%."
    if choice.payment == "custom":
        return _custom_text(choice)
    if choice.payment == "deferral":
        days = _days(choice.payment_days, 45)
        return f"Отсрочка платежа — {calendar_days_phrase(days)}."
    if choice.payment == "split_50_50":
        first, second = payment_tranche_kopecks("split_50_50", total)
        days = _days(choice.payment_days, 3)
        due = _require_date(choice.payment_date, "Укажите дату первого аванса")
        return (
            f"{money_phrase(first)} — предварительная оплата в размере 50% "
            f"до {format_date(due)} включительно. "
            f"{money_phrase(second)} — за {days} ({integer_in_words(days)}) {day_unit(days)} "
            "до планируемого начала отгрузки."
        )
    first, second, rest = payment_tranche_kopecks(
        "split_share",
        total,
        share_percent=choice.second_share_percent,
    )
    days = _days(choice.payment_days, 5)
    first_date = _require_date(choice.payment_date, "Укажите дату первого аванса")
    second_date = _require_date(choice.second_payment_date, "Укажите дату второго аванса")
    share = int(choice.second_share_percent or 0)
    return (
        f"{money_phrase(first)} — предварительная оплата в размере 50% "
        f"до {format_date(first_date)} включительно. "
        f"{money_phrase(second)} — предварительная оплата в размере {share}% "
        f"до {format_date(second_date)} включительно. "
        f"{money_phrase(rest)} — остаток в течение {calendar_days_phrase(days)} "
        "после поставки партии. Партия — товар по одному УПД."
    )


def _term_paragraph(choice: SpecificationChoice, *, has_piles: bool) -> str:
    if choice.term == "by_date":
        due = _require_date(choice.term_date, "Укажите дату поставки")
        sentence = f"Поставщик обязуется поставить товар не позднее {date_in_words(due)}"
        clause = _prepay_clause(choice.payment)
        if clause:
            return f"{sentence} {clause}."
        return f"{sentence}."
    if choice.term == "pile_rhythm":
        if not has_piles:
            raise SpecificationTextError("Ритм свай доступен только если в КП есть сваи")
        start = _require_date(choice.term_date, "Укажите дату начала поставок")
        count = _positive_int(choice.pile_count, "Укажите число свай")
        unit = "неделю" if choice.pile_unit == "week" else "день"
        return f"Начало поставок с {format_date(start)}. По {count} свай в {unit}."
    days = _days(choice.term_days, 10)
    return (
        "Поставщик обязуется поставить товар в течение "
        f"{calendar_days_phrase(days)} со дня, следующего за днем поступления "
        "денежных средств на расчетный счет поставщика в полном объеме."
    )


def _delivery_paragraph(choice: SpecificationChoice) -> str:
    if choice.delivery == "pickup":
        return f"Самовывоз со склада готовой продукции, {PICKUP_ADDRESS}."
    address = _site_address(choice)
    return (
        f"Поставка на объект покупателя по адресу: {address}. "
        "Стоимость доставки включена в стоимость товара."
    )


def _prepay_clause(payment: PaymentKind) -> str | None:
    if payment in ("prepay_100", "custom"):
        return PREPAY_CLAUSE_FULL
    if payment in ("split_50_50", "split_share"):
        return PREPAY_CLAUSE_PARTIAL
    return None


def _validate(choice: SpecificationChoice, *, has_piles: bool) -> None:
    if choice.payment not in ("prepay_100", "split_50_50", "split_share", "deferral", "custom"):
        raise SpecificationTextError("Выберите схему оплаты")
    if choice.term not in ("by_date", "pile_rhythm", "after_payment"):
        raise SpecificationTextError("Выберите срок поставки")
    if choice.delivery not in ("pickup", "site"):
        raise SpecificationTextError("Выберите вид поставки")
    if choice.payment == "split_50_50":
        _require_date(choice.payment_date, "Укажите дату первого аванса")
        _days(choice.payment_days, 3)
    elif choice.payment == "split_share":
        if choice.second_share_percent not in _SHARE_PERCENTS:
            raise SpecificationTextError("Доля второго аванса — 10, 20, 30 или 40%")
        _require_date(choice.payment_date, "Укажите дату первого аванса")
        _require_date(choice.second_payment_date, "Укажите дату второго аванса")
        _days(choice.payment_days, 5)
    elif choice.payment == "deferral":
        _days(choice.payment_days, 45)
    elif choice.payment == "custom":
        _custom_text(choice)
    if choice.term == "by_date":
        _require_date(choice.term_date, "Укажите дату поставки")
    elif choice.term == "pile_rhythm":
        if not has_piles:
            raise SpecificationTextError("Ритм свай доступен только если в КП есть сваи")
        _require_date(choice.term_date, "Укажите дату начала поставок")
        _positive_int(choice.pile_count, "Укажите число свай")
        if choice.pile_unit not in ("week", "day"):
            raise SpecificationTextError("Выберите ритм: в неделю или в день")
    else:
        _days(choice.term_days, 10)
    if choice.delivery == "site":
        _site_address(choice)


def _custom_text(choice: SpecificationChoice) -> str:
    text = (choice.custom_text or "").strip()
    if not text:
        raise SpecificationTextError("Свой текст не может быть пустым")
    if len(text) > 2000:
        raise SpecificationTextError("Свой текст не длиннее 2000 символов")
    return text


def _site_address(choice: SpecificationChoice) -> str:
    address = (choice.delivery_address or "").strip()
    if not address:
        raise SpecificationTextError("Укажите адрес объекта")
    if len(address) > 300:
        raise SpecificationTextError("Адрес объекта не длиннее 300 символов")
    return address


def _days(value: int | None, default: int) -> int:
    days = default if value is None else value
    if isinstance(days, bool) or not isinstance(days, int) or days < 1 or days > 365:
        raise SpecificationTextError("Число дней — целое от 1 до 365")
    return days


def _require_date(value: date | None, message: str) -> date:
    if not isinstance(value, date):
        raise SpecificationTextError(message)
    return value


def _positive_int(value: int | None, message: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise SpecificationTextError(message)
    return value


def _percent_kopecks(total: int, percent: int) -> int:
    amount = (Decimal(total) * Decimal(percent) / Decimal(100)).quantize(
        _KOPECK, rounding=ROUND_HALF_UP
    )
    return int(amount)


def _triplet(value: int, *, feminine: bool) -> str:
    ones = _ONES_F if feminine else _ONES_M
    parts: list[str] = []
    hundreds = value // 100
    if hundreds:
        parts.append(_HUNDREDS[hundreds])
    rest = value % 100
    if 10 <= rest <= 19:
        parts.append(_TEENS[rest - 10])
    else:
        tens = rest // 10
        unit = rest % 10
        if tens:
            parts.append(_TENS[tens])
        if unit:
            parts.append(ones[unit])
    return " ".join(parts)


def _plural(value: int, one: str, few: str, many: str) -> str:
    number = abs(value) % 100
    if 11 <= number <= 14:
        return many
    last = number % 10
    if last == 1:
        return one
    if 2 <= last <= 4:
        return few
    return many


def specification_document(
    choice: SpecificationChoice,
    *,
    spec_date: str,
    composition_hash: str,
    has_piles: bool,
) -> dict[str, Any]:
    """JSON колонки. Дата первого сохранения передаётся снаружи."""
    _validate(choice, has_piles=has_piles)
    return {
        "payment": choice.payment,
        "payment_date": _iso(choice.payment_date),
        "payment_days": _stored_payment_days(choice),
        "second_share_percent": (
            choice.second_share_percent if choice.payment == "split_share" else None
        ),
        "second_payment_date": (
            _iso(choice.second_payment_date) if choice.payment == "split_share" else None
        ),
        "custom_text": _custom_text(choice) if choice.payment == "custom" else None,
        "term": choice.term,
        "term_date": _iso(choice.term_date) if choice.term in ("by_date", "pile_rhythm") else None,
        "pile_count": choice.pile_count if choice.term == "pile_rhythm" else None,
        "pile_unit": choice.pile_unit if choice.term == "pile_rhythm" else None,
        "term_days": _days(choice.term_days, 10) if choice.term == "after_payment" else None,
        "delivery": choice.delivery,
        "delivery_address": _site_address(choice) if choice.delivery == "site" else None,
        "spec_date": spec_date,
        "composition_hash": composition_hash,
    }


def choice_from_document(document: Mapping[str, Any]) -> SpecificationChoice:
    try:
        return SpecificationChoice(
            payment=document["payment"],  # type: ignore[arg-type]
            term=document["term"],  # type: ignore[arg-type]
            delivery=document["delivery"],  # type: ignore[arg-type]
            payment_date=_date_from_iso(document.get("payment_date")),
            payment_days=_optional_int(document.get("payment_days")),
            second_share_percent=_optional_int(document.get("second_share_percent")),
            second_payment_date=_date_from_iso(document.get("second_payment_date")),
            custom_text=None if document.get("custom_text") is None else str(document.get("custom_text")),
            term_date=_date_from_iso(document.get("term_date")),
            pile_count=_optional_int(document.get("pile_count")),
            pile_unit=document.get("pile_unit"),  # type: ignore[arg-type]
            term_days=_optional_int(document.get("term_days")),
            delivery_address=(
                None
                if document.get("delivery_address") is None
                else str(document.get("delivery_address"))
            ),
        )
    except (KeyError, TypeError, ValueError) as exc:
        raise SpecificationTextError("Спецификация сохранена некорректно") from exc


def _stored_payment_days(choice: SpecificationChoice) -> int | None:
    if choice.payment == "split_50_50":
        return _days(choice.payment_days, 3)
    if choice.payment == "split_share":
        return _days(choice.payment_days, 5)
    if choice.payment == "deferral":
        return _days(choice.payment_days, 45)
    return None


def _iso(value: date | None) -> str | None:
    if value is None:
        return None
    return value.isoformat()


def _date_from_iso(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


def _optional_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool) or not isinstance(value, int):
        raise SpecificationTextError("Число дней — целое от 1 до 365")
    return value
