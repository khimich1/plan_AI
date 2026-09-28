"""Картинка карточки: препроцесс, Extract, Verify. Промпт полей договора, не марок."""

from __future__ import annotations

import base64
import json
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Protocol

from core.config.settings import Settings, get_settings
from core.ocr.image_preprocess import preprocess_image_for_ocr
from core.ocr.prompts import (
    CONTRACT_CARD_EXTRACT_PROMPT,
    get_contract_card_holes_prompt,
    get_contract_card_reread_prompt,
)
from core.ocr.providers.openai import image_mime_type, load_image_payload
from core.supply_contract import normalize_digits, normalize_email
from core.supply_contract_card_parse import (
    FIELD_NAMES,
    CardParseResult,
    _email,
    _suspicious_address,
    _suspicious_bank,
    accept_account,
    accept_bik,
    accept_bundle,
    accept_corr_account,
    accept_inn,
    accept_kpp,
    accept_ogrn,
    card_holes,
    empty_fields,
    frozen_fields,
    parse_card_text,
)

_DIGIT_FIELDS = ("inn", "kpp", "ogrn", "account", "corr_account", "bik")
_CHOICE_FIELDS = ("legal_form", "signatory_verb", "authority_basis")
_LEGAL_FORMS = frozenset({"ooo", "ao", "ip", "kfh", "person"})


class ContractCardProvider(Protocol):
    async def extract_contract_fields(
        self,
        *,
        user_text: str,
        image_base64: str,
        mime_type: str,
    ) -> tuple[dict[str, Any], float]:
        """Первый проход: поля карточки."""

    async def reread_contract_fields(
        self,
        *,
        user_text: str,
        image_base64: str,
        mime_type: str,
    ) -> tuple[dict[str, Any], float]:
        """Один повтор только по именам полей, которые контроль не принял."""


def apply_contract_verify(
    extract_fields: dict[str, Any],
    verify_result: dict[str, Any] | None,
) -> tuple[dict[str, Any], bool]:
    """Пустой Verify не затирает Extract и помечает сверку неуспешной.

    Непустые corrections подменяют поля. Пустые corrections оставляют Extract.
    """
    result = verify_result or {}
    verified = result.get("fields")
    corrections = result.get("corrections") or []
    if not isinstance(verified, dict) or not _nonempty_fields(verified):
        return dict(extract_fields), True
    if not corrections:
        return dict(extract_fields), False
    return _coerce_fields(verified), False


def fields_from_model(raw: dict[str, Any] | None) -> dict[str, Any]:
    return _coerce_fields(raw or {})


async def recognize_contract_card(
    data: bytes,
    *,
    filename: str,
    provider: ContractCardProvider,
    settings: Settings | None = None,
) -> CardParseResult:
    cfg = settings or get_settings()
    suffix = Path(filename).suffix or ".bin"
    with NamedTemporaryFile(suffix=suffix) as handle:
        handle.write(data)
        handle.flush()
        image_base64, mime_type = _prepared_image(
            handle.name,
            data,
            min_short_side=cfg.ocr_verify_auto_min_short_side,
        )
        extracted, _cost = await provider.extract_contract_fields(
            user_text=CONTRACT_CARD_EXTRACT_PROMPT,
            image_base64=image_base64,
            mime_type=mime_type,
        )
        controlled = _control_digits(fields_from_model(extracted))
        missing = _missing_digits(controlled)
        if missing:
            reread, _reread_cost = await provider.reread_contract_fields(
                user_text=get_contract_card_reread_prompt(missing),
                image_base64=image_base64,
                mime_type=mime_type,
            )
            controlled = _fill_missing_digits(controlled, fields_from_model(reread), missing)
    parsed = parse_card_text(_fields_as_text(controlled))
    for name in _DIGIT_FIELDS:
        parsed.fields[name] = controlled.get(name)
    for name, value in controlled.items():
        if name in _DIGIT_FIELDS or not value:
            continue
        if name in _CHOICE_FIELDS:
            choice = _canonical_choice(name, str(value))
            if choice:
                parsed.fields[name] = choice
            continue
        parsed.fields[name] = value
    parsed.verify_failed = False
    parsed.doubtful = _doubtful_for(parsed.fields)
    parsed.source_text = ""
    parsed.banks = []
    parsed.bank_unpaired = False
    parsed.inn_disputed = False
    return parsed


def _prepared_image(path: str, original: bytes, *, min_short_side: int) -> tuple[str, str]:
    processed = preprocess_image_for_ocr(path, min_short_side=min_short_side)
    if processed is not None and processed.applied:
        encoded = base64.b64encode(processed.image_data).decode()
        return encoded, processed.mime_type
    try:
        _raw, encoded, mime = load_image_payload(path)
        return encoded, mime
    except OSError:
        return base64.b64encode(original).decode(), image_mime_type(path, original)


def _canonical_choice(name: str, value: str) -> str | None:
    """Слово с карточки → пункт списка. Чужое слово в поле не кладётся."""
    text = value.strip().lower().replace("ё", "е")
    if name == "legal_form":
        return text if text in _LEGAL_FORMS else None
    if name == "signatory_verb":
        if "действующей" in text:
            return "действующей"
        if "действующего" in text:
            return "действующего"
        return None
    if name == "authority_basis":
        if "доверен" in text:
            return "доверенность"
        if "устав" in text:
            return "устав"
        return None
    return None


def _nonempty_fields(fields: dict[str, Any]) -> bool:
    return any(str(value).strip() for value in fields.values() if value is not None)


def _coerce_fields(raw: dict[str, Any]) -> dict[str, Any]:
    fields = empty_fields()
    nested = raw.get("fields") if isinstance(raw.get("fields"), dict) else raw
    for name in FIELD_NAMES:
        value = nested.get(name)
        if value is None:
            continue
        text = str(value).strip()
        fields[name] = text or None
    return fields


def _control_digits(fields: dict[str, Any]) -> dict[str, Any]:
    controlled = _normalize_known(fields)
    controlled["inn"] = accept_inn(controlled.get("inn"))
    controlled["kpp"] = accept_kpp(controlled.get("kpp"))
    controlled["ogrn"] = accept_ogrn(controlled.get("ogrn"))
    controlled["bik"] = accept_bik(controlled.get("bik"))
    controlled["corr_account"] = accept_corr_account(controlled.get("corr_account"))
    controlled["account"] = accept_account(controlled.get("account"), controlled.get("bik"))
    return controlled


def _missing_digits(fields: dict[str, Any]) -> list[str]:
    return [name for name in _DIGIT_FIELDS if not fields.get(name)]


def _fill_missing_digits(
    fields: dict[str, Any],
    extra: dict[str, Any],
    missing: list[str],
) -> dict[str, Any]:
    filled = dict(fields)
    pending = [name for name in missing if not filled.get(name)]
    if "bik" in pending:
        filled["bik"] = accept_bik(extra.get("bik"))
    for name in pending:
        if filled.get(name):
            continue
        if name == "account":
            filled["account"] = accept_account(extra.get("account"), filled.get("bik"))
        elif name == "corr_account":
            filled["corr_account"] = accept_corr_account(extra.get("corr_account"))
        elif name == "inn":
            filled["inn"] = accept_inn(extra.get("inn"))
        elif name == "kpp":
            filled["kpp"] = accept_kpp(extra.get("kpp"))
        elif name == "ogrn":
            filled["ogrn"] = accept_ogrn(extra.get("ogrn"))
    return filled


def _normalize_known(fields: dict[str, Any]) -> dict[str, Any]:
    normalized = dict(fields)
    for key in _DIGIT_FIELDS:
        if normalized.get(key):
            normalized[key] = normalize_digits(str(normalized[key])) or None
    if normalized.get("email"):
        normalized["email"] = normalize_email(str(normalized["email"])) or None
    return normalized


def _fields_as_text(fields: dict[str, Any]) -> str:
    lines = []
    labels = {
        "inn": "ИНН",
        "kpp": "КПП",
        "ogrn": "ОГРН",
        "bik": "БИК",
        "account": "Расчётный счёт",
        "corr_account": "Корреспондентский счёт",
        "email": "E-mail",
        "legal_address": "Юридический адрес",
        "bank_name": "Банк",
        "signatory_name": "Директор",
        "signatory_position": "",
    }
    for key, label in labels.items():
        value = fields.get(key)
        if not value:
            continue
        if key == "signatory_position":
            continue
        if key == "signatory_name" and fields.get("signatory_position"):
            lines.append(f"{fields['signatory_position']} {value}")
        else:
            lines.append(f"{label} {value}" if label else str(value))
    if fields.get("short_name"):
        lines.append(f"«{fields['short_name']}»")
    if fields.get("legal_form") == "ooo":
        lines.append("ООО")
    return "\n".join(lines)


def _doubtful_for(fields: dict[str, Any]) -> list[str]:
    return list(parse_card_text(_fields_as_text(fields)).doubtful)


def draft_json(fields: dict[str, Any]) -> str:
    return json.dumps(fields, ensure_ascii=False)


_MODEL_DOUBTFUL = ("legal_address", "signatory_position", "signatory_name")
_DIGIT_ACCEPT = {
    "inn": accept_inn,
    "kpp": accept_kpp,
    "ogrn": accept_ogrn,
    "bik": accept_bik,
    "corr_account": accept_corr_account,
}


class CardTextProvider(Protocol):
    async def complete_text_json(self, *, user_text: str) -> tuple[dict[str, Any], float]:
        """Один текстовый чат без картинки."""


async def fill_card_holes(
    parsed: CardParseResult,
    text: str,
    provider: CardTextProvider,
) -> CardParseResult:
    """Дописывает только дыры. Чистая карточка провайдера не вызывает."""
    holes = card_holes(parsed)
    if not holes:
        return parsed
    prompt = get_contract_card_holes_prompt(
        holes=holes,
        frozen=frozen_fields(parsed),
        card_text=(parsed.source_text or text)[:20_000],
    )
    try:
        raw, _cost = await provider.complete_text_json(user_text=prompt)
    except Exception:
        return parsed
    if not isinstance(raw, dict):
        return parsed
    return _apply_hole_payload(parsed, _payload_fields(raw), holes)


def _payload_fields(raw: dict[str, Any]) -> dict[str, Any]:
    nested = raw.get("fields")
    if not isinstance(nested, dict):
        return raw
    merged = dict(nested)
    for key, value in raw.items():
        if key != "fields":
            merged[key] = value
    return merged


def _apply_hole_payload(
    parsed: CardParseResult,
    payload: dict[str, Any],
    holes: list[str],
) -> CardParseResult:
    frozen = set(frozen_fields(parsed))
    if "banks" in holes:
        _apply_model_banks(parsed, payload.get("banks"))
    order = [key for key in ("bik", "inn", "kpp", "ogrn", "corr_account", "account") if key in holes]
    order.extend(key for key in holes if key not in order and key != "banks")
    for key in order:
        if key in frozen or key not in payload:
            continue
        _apply_model_field(parsed, key, payload.get(key))
    if parsed.inn_disputed:
        _mark_parsed(parsed, "inn")
    for key in _MODEL_DOUBTFUL:
        if key in holes and parsed.fields.get(key):
            _mark_parsed(parsed, key)
    return parsed


def _apply_model_banks(parsed: CardParseResult, raw: Any) -> None:
    if not isinstance(raw, list):
        return
    bundles = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        bundle = accept_bundle(
            item.get("bank_name"),
            str(item.get("account") or ""),
            item.get("corr_account"),
            str(item.get("bik") or ""),
        )
        if bundle:
            bundles.append(bundle)
    if len(bundles) >= 2:
        parsed.banks = bundles
        parsed.bank_unpaired = False
        for key in ("bank_name", "account", "corr_account", "bik"):
            parsed.fields[key] = None
        return
    if len(bundles) != 1:
        return
    frozen = frozen_fields(parsed)
    for key, value in bundles[0].items():
        if key in frozen or not value:
            continue
        parsed.fields[key] = value


def _apply_model_field(parsed: CardParseResult, key: str, value: Any) -> None:
    if value is None:
        return
    if key == "account":
        accepted = accept_account(str(value), parsed.fields.get("bik"))
        if accepted is None:
            return
        parsed.fields["account"] = accepted
        _unmark(parsed, "account")
        return
    accept = _DIGIT_ACCEPT.get(key)
    if accept is not None:
        accepted = accept(str(value))
        if accepted is None:
            return
        parsed.fields[key] = accepted
        if not (key == "inn" and parsed.inn_disputed):
            _unmark(parsed, key)
        return
    if key == "legal_form":
        choice = _canonical_choice("legal_form", str(value))
        if choice:
            parsed.fields["legal_form"] = choice
            _unmark(parsed, "legal_form")
        return
    text = str(value).strip()
    if not text:
        return
    if key == "legal_address" and _suspicious_address(text):
        return
    if key == "bank_name" and _suspicious_bank(text):
        return
    if key == "email":
        email = _email(text)
        if not email:
            return
        parsed.fields["email"] = email
        _unmark(parsed, "email")
        return
    parsed.fields[key] = text
    if key not in _MODEL_DOUBTFUL:
        _unmark(parsed, key)


def _mark_parsed(parsed: CardParseResult, key: str) -> None:
    if key not in parsed.doubtful:
        parsed.doubtful.append(key)


def _unmark(parsed: CardParseResult, key: str) -> None:
    parsed.doubtful = [item for item in parsed.doubtful if item != key]
