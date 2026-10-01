"""Выдача и чтение действующего договора поставки по КП архива."""

from __future__ import annotations

import logging
import shutil
import subprocess
import tempfile
import uuid
from datetime import date
from pathlib import Path
from typing import Any

from app.repositories.kp_offers_repository import KpOffersRepository
from app.repositories.supply_contract_repository import SupplyContractRepository
from app.schemas.supply_contract import SupplyContractCreate
from app.security.offer_access import assert_offer_write_access
from core.config.settings import PROJECT_ROOT
from app.services.supply_contract_import import (
    SupplyContractImportError,
    parse_lawyer_sheet,
    suggest_names,
)
from core.supply_contract import (
    CANCELLED_STATUS,
    CONTRACT_STATUSES,
    LEGAL_FORMS_REQUIRING_KPP,
    LEGAL_FORMS_REQUIRING_POSITION,
    MSG_BIND_COUNTERPARTY,
    MSG_NOT_ARCHIVED,
    SCAN_NOTES,
    attach_contract_number as stamp_contract_number,
    duplicate_contract_message,
    normalize_digits,
    normalize_email,
)

_SCANS_DIR = PROJECT_ROOT / ".app_data" / "supply-contract-scans"
from core.supply_contract_card_parse import (
    card_holes,
    extract_docx_text,
    extract_pdf_text,
    extract_xlsx_text,
    has_text_layer,
    is_docx,
    is_pdf,
    is_xlsx,
    parse_card_text,
)
from core.edo_agreement_docx import render_edo_agreement_docx
from core.supply_contract_docx import render_supply_contract_docx

MSG_UNSUPPORTED_CARD = "Неподдерживаемый формат файла. Допустимы DOCX, PDF, JPEG или PNG."
MSG_NOT_A_CARD = "это не карточка контрагента"
MSG_CONTRACT_MISSING = "Договор ещё не оформлен"
MSG_EDO_REQUISITES = "В договоре не хватает реквизитов покупателя для соглашения об ЭДО"
MSG_LIBREOFFICE = "Не удалось собрать PDF. На сервере недоступен LibreOffice."
_CARD_REQUISITES = ("inn", "ogrn", "account", "bik")
_EDO_REQUIRED = (
    "legal_form",
    "full_name",
    "signatory_name",
    "inn",
    "ogrn",
    "legal_address",
    "email",
)
_PDF_TIMEOUT_SECONDS = 60
_DOWNLOAD_FORMATS = frozenset({"docx", "pdf"})
_log = logging.getLogger(__name__)


class SupplyContractError(Exception):
    pass


class SupplyContractNotFoundError(SupplyContractError):
    pass


class SupplyContractValidationError(SupplyContractError):
    """Гейт архива: нет контрагента или КП не в архиве."""


class SupplyContractFieldError(SupplyContractError):
    """Поля покупателя не проходят правила вида."""


class SupplyContractPdfError(SupplyContractError):
    """Нет LibreOffice или конвертация docx в pdf не удалась."""


class SupplyContractService:
    def __init__(
        self,
        *,
        db_path: str,
        today: date | None = None,
        card_recognizer: Any = None,
        card_text_provider: Any = None,
        scans_dir: Path | None = None,
    ) -> None:
        self.repository = SupplyContractRepository(db_path=db_path)
        self.offers = KpOffersRepository(db_path)
        self._today = today
        self.card_recognizer = card_recognizer
        self.card_text_provider = card_text_provider
        self.scans_dir = scans_dir or _SCANS_DIR

    def get_for_kp(self, kp_id: int, *, user: dict) -> dict[str, Any] | None:
        raw = self._offer_for_read(kp_id, user)
        counterparty_id = raw.get("counterparty_id")
        if not counterparty_id:
            raise SupplyContractValidationError(MSG_BIND_COUNTERPARTY)
        row = self.repository.get_active_by_counterparty(int(counterparty_id))
        return _present(row) if row else None

    def create_for_kp(
        self,
        kp_id: int,
        payload: SupplyContractCreate,
        *,
        user: dict,
    ) -> dict[str, Any]:
        raw = self._offer_for_write(kp_id, user)
        counterparty_id = raw.get("counterparty_id")
        if not counterparty_id:
            raise SupplyContractValidationError(MSG_BIND_COUNTERPARTY)
        existing = self.repository.get_active_by_counterparty(int(counterparty_id))
        if existing is not None:
            return _present(existing)
        fields = self._validated_fields(payload)
        contract_date = payload.contract_date or self._today or date.today()
        manager_name = str(raw.get("manager_name") or "").strip() or "—"
        created, _inserted = self.repository.create_active(
            fields,
            counterparty_id=int(counterparty_id),
            manager_name=manager_name,
            contract_date=contract_date,
            user_id=int(user["id"]),
        )
        return _present(created)

    def attach_contract_number(
        self,
        document: dict[str, Any],
        counterparty_id: int,
    ) -> dict[str, Any]:
        """Шапка JSON. Номер — действующий договор контрагента. Отправки нет."""
        row = self.repository.get_active_by_counterparty(int(counterparty_id))
        return stamp_contract_number(document, int(counterparty_id), row)

    def import_lawyer_sheet(self, data: bytes, *, user: dict) -> list[dict[str, Any]]:
        if user.get("role") != "admin":
            raise SupplyContractValidationError("Импорт листа доступен только администратору")
        try:
            parsed = parse_lawyer_sheet(data)
            stored = self.repository.insert_imported(parsed, user_id=int(user["id"]))
        except SupplyContractImportError as exc:
            raise SupplyContractValidationError(str(exc)) from exc
        except ValueError as exc:
            text = str(exc)
            if text.startswith("duplicate:"):
                raise SupplyContractValidationError(
                    f"Номер {text.split(':', 1)[1]} уже есть в реестре"
                ) from exc
            raise
        parties = self.repository.list_counterparty_names()
        return [_imported_view(row, parties) for row in stored]

    def link_imported_contract(
        self,
        contract_id: int,
        counterparty_id: int,
        *,
        user: dict,
    ) -> dict[str, Any]:
        self._assert_writer(user)
        try:
            updated = self.repository.link_counterparty(
                contract_id,
                int(counterparty_id),
                user_id=int(user["id"]),
            )
        except ValueError as exc:
            text = str(exc)
            if text == "counterparty":
                raise SupplyContractValidationError("Контрагент не найден") from exc
            if text == "already-linked":
                raise SupplyContractValidationError(
                    "Договор уже связан с другим контрагентом"
                ) from exc
            raise SupplyContractValidationError(duplicate_contract_message(text)) from exc
        if updated is None:
            raise SupplyContractNotFoundError(f"Договор №{contract_id} не найден")
        return _present(updated)

    def list_registry(self, *, user: dict) -> list[dict[str, Any]]:
        self._assert_writer(user)
        return self.repository.list_registry()

    def update_contract(
        self,
        contract_id: int,
        *,
        user: dict,
        status: str | None = None,
        scan_note: str | None = None,
        scan_note_set: bool = False,
        contract_date: date | str | None = None,
        scan_bytes: bytes | None = None,
        scan_filename: str = "",
    ) -> dict[str, Any]:
        self._assert_writer(user)
        if status is not None and status not in CONTRACT_STATUSES:
            raise SupplyContractFieldError("Неизвестный статус договора")
        if scan_note_set:
            note = (scan_note or "").strip() or None
            if note is not None and note not in SCAN_NOTES:
                raise SupplyContractFieldError("Отметка скана: «получен скан» или «внесены правки»")
            scan_note = note
        parsed_date = _parse_date(contract_date) if contract_date else None
        current = self.repository.get_by_id(contract_id)
        if current is None:
            raise SupplyContractNotFoundError(f"Договор №{contract_id} не найден")
        new_path, previous = self._store_scan(contract_id, current, scan_bytes, scan_filename)
        try:
            updated = self.repository.apply_update(
                contract_id,
                user_id=int(user["id"]),
                status=status,
                scan_note=scan_note,
                scan_note_set=scan_note_set,
                contract_date=parsed_date.isoformat() if parsed_date else None,
                scan_path=new_path,
                scan_replaced=new_path is not None,
            )
        except Exception:
            if new_path:
                Path(new_path).unlink(missing_ok=True)
            raise
        if updated is None:
            raise SupplyContractNotFoundError(f"Договор №{contract_id} не найден")
        if new_path and previous:
            _delete_owned_scan(self.scans_dir, previous)
        return _present(updated)

    def replace_contract(
        self,
        contract_id: int,
        payload: SupplyContractCreate,
        *,
        user: dict,
    ) -> dict[str, Any]:
        self._assert_writer(user)
        current = self.repository.get_by_id(contract_id)
        if current is None:
            raise SupplyContractNotFoundError(f"Договор №{contract_id} не найден")
        if current["status"] != CANCELLED_STATUS:
            raise SupplyContractValidationError(duplicate_contract_message(str(current["number"])))
        counterparty_id = current.get("counterparty_id")
        if not counterparty_id:
            raise SupplyContractValidationError("Сначала свяжите договор с контрагентом")
        active = self.repository.get_active_by_counterparty(int(counterparty_id))
        if active is not None:
            raise SupplyContractValidationError(duplicate_contract_message(str(active["number"])))
        fields = self._validated_fields(payload)
        contract_date = payload.contract_date or self._today or date.today()
        created, inserted = self.repository.create_active(
            fields,
            counterparty_id=int(counterparty_id),
            manager_name=str(current.get("manager_name") or "—"),
            contract_date=contract_date,
            user_id=int(user["id"]),
        )
        if not inserted:
            raise SupplyContractValidationError(duplicate_contract_message(str(created["number"])))
        return _present(created)

    def document_for_kp(
        self,
        kp_id: int,
        *,
        user: dict,
        file_format: str = "docx",
    ) -> tuple[bytes, str]:
        contract = self.get_for_kp(kp_id, user=user)
        if contract is None:
            raise SupplyContractNotFoundError(MSG_CONTRACT_MISSING)
        payload = render_supply_contract_docx(contract)
        return _as_download(payload, str(contract["number"]), "dogovor", file_format)

    def edo_agreement_for_kp(
        self,
        kp_id: int,
        *,
        user: dict,
        file_format: str = "docx",
    ) -> tuple[bytes, str]:
        contract = self.get_for_kp(kp_id, user=user)
        if contract is None:
            raise SupplyContractNotFoundError(MSG_CONTRACT_MISSING)
        _require_edo_buyer(contract)
        payload = render_edo_agreement_docx(contract)
        return _as_download(payload, str(contract["number"]), "soglashenie-edo", file_format)

    async def parse_for_kp(
        self,
        kp_id: int,
        data: bytes,
        *,
        filename: str,
        user: dict,
    ) -> dict[str, Any]:
        """Поля карточки без строки договора и без выдачи номера."""
        raw = self._offer_for_write(kp_id, user)
        if not raw.get("counterparty_id"):
            raise SupplyContractValidationError(MSG_BIND_COUNTERPARTY)
        return await self._parse_upload(data, filename=filename)

    async def _parse_upload(self, data: bytes, *, filename: str) -> dict[str, Any]:
        if not data:
            raise SupplyContractValidationError("Пустой файл.")
        parsed = await self._read_card(data, filename=filename)
        _require_counterparty_card(parsed)
        return parsed

    async def _read_card(self, data: bytes, *, filename: str) -> dict[str, Any]:
        if is_docx(data):
            return await self._read_text(extract_docx_text(data))
        if is_xlsx(data):
            return await self._read_text(extract_xlsx_text(data))
        if is_pdf(data):
            text = extract_pdf_text(data)
            if has_text_layer(text):
                return await self._read_text(text)
            return await self._recognize_image(data, filename=filename or "scan.pdf")
        if _is_image(data):
            return await self._recognize_image(data, filename=filename or "card.jpg")
        raise SupplyContractValidationError(MSG_UNSUPPORTED_CARD)

    async def _read_text(self, text: str) -> dict[str, Any]:
        parsed = parse_card_text(text)
        if card_holes(parsed):
            parsed = await self._fill_text_holes(parsed, text)
        return parsed.as_dict()

    async def _fill_text_holes(self, parsed: Any, text: str) -> Any:
        provider = self.card_text_provider
        if provider is None:
            if self.card_recognizer is not None:
                return parsed
            provider = _default_text_provider()
        from core.supply_contract_card_ocr import fill_card_holes

        return await fill_card_holes(parsed, text, provider)

    async def _recognize_image(self, data: bytes, *, filename: str) -> dict[str, Any]:
        recognizer = self.card_recognizer
        if recognizer is None:
            from core.supply_contract_card_ocr import recognize_contract_card

            async def recognizer(payload: bytes, *, filename: str) -> Any:
                provider = _default_card_provider()
                return await recognize_contract_card(
                    payload,
                    filename=filename,
                    provider=provider,
                )

        result = await recognizer(data, filename=filename)
        if hasattr(result, "as_dict"):
            return result.as_dict()
        return result

    def update_contract_date(
        self,
        contract_id: int,
        contract_date: date | str,
        *,
        user: dict,
    ) -> dict[str, Any]:
        parsed = (
            contract_date
            if isinstance(contract_date, date)
            else date.fromisoformat(contract_date)
        )
        updated = self.repository.update_contract_date(
            contract_id,
            parsed,
            user_id=int(user["id"]),
        )
        if updated is None:
            raise SupplyContractNotFoundError(f"Договор №{contract_id} не найден")
        return _present(updated)

    def _assert_writer(self, user: dict) -> None:
        assert_offer_write_access(user, {})

    def _store_scan(
        self,
        contract_id: int,
        current: dict[str, Any],
        data: bytes | None,
        filename: str,
    ) -> tuple[str | None, str | None]:
        if not data:
            return None, None
        suffix = _scan_suffix(data, filename)
        directory = self.scans_dir / str(int(contract_id))
        directory.mkdir(parents=True, exist_ok=True)
        stored = directory / f"{uuid.uuid4().hex}{suffix}"
        stored.write_bytes(data)
        previous = current.get("scan_path")
        return str(stored), str(previous) if previous else None

    def _offer_for_read(self, kp_id: int, user: dict) -> dict[str, Any]:
        raw = self._load_offer(kp_id, user)
        if raw.get("status") not in {"в архиве", "на согласовании"}:
            raise SupplyContractValidationError(MSG_NOT_ARCHIVED)
        return raw

    def _offer_for_write(self, kp_id: int, user: dict) -> dict[str, Any]:
        raw = self._load_offer(kp_id, user)
        if raw.get("status") != "в архиве":
            raise SupplyContractValidationError(MSG_NOT_ARCHIVED)
        return raw

    def _load_offer(self, kp_id: int, user: dict) -> dict[str, Any]:
        raw = self.offers.get_by_id(kp_id)
        if not raw:
            raise SupplyContractNotFoundError(f"КП №{kp_id} не найдено")
        assert_offer_write_access(user, raw)
        return raw

    def _validated_fields(self, payload: SupplyContractCreate) -> dict[str, Any]:
        kpp = normalize_digits(payload.kpp) or None
        if payload.legal_form in LEGAL_FORMS_REQUIRING_KPP and not kpp:
            raise SupplyContractFieldError("Для ООО и АО укажите КПП")
        position = (payload.signatory_position or "").strip() or None
        if payload.legal_form in LEGAL_FORMS_REQUIRING_POSITION and not position:
            raise SupplyContractFieldError("Укажите должность подписанта")
        if payload.authority_basis == "доверенность":
            if not (payload.poa_number or "").strip() or not (payload.poa_date or "").strip():
                raise SupplyContractFieldError("Укажите номер и дату доверенности")
        inn = normalize_digits(payload.inn)
        ogrn = normalize_digits(payload.ogrn)
        account = normalize_digits(payload.account)
        corr_account = normalize_digits(payload.corr_account)
        bik = normalize_digits(payload.bik)
        email = normalize_email(payload.email)
        if not inn or not ogrn or not account or not corr_account or not bik or not email:
            raise SupplyContractFieldError("Проверьте ИНН, ОГРН, счета, БИК и почту")
        return {
            "legal_form": payload.legal_form,
            "full_name": payload.full_name.strip(),
            "short_name": payload.short_name.strip(),
            "signatory_position": position,
            "signatory_name": payload.signatory_name.strip(),
            "signatory_verb": payload.signatory_verb,
            "authority_basis": payload.authority_basis,
            "poa_number": (payload.poa_number or "").strip() or None,
            "poa_date": (payload.poa_date or "").strip() or None,
            "inn": inn,
            "kpp": kpp,
            "ogrn": ogrn,
            "legal_address": payload.legal_address.strip(),
            "postal_address": (payload.postal_address or "").strip() or None,
            "phone": (payload.phone or "").strip() or None,
            "email": email,
            "bank_name": payload.bank_name.strip(),
            "account": account,
            "corr_account": corr_account,
            "bik": bik,
            "edo_operator": (payload.edo_operator or "").strip() or None,
            "edo_id": (payload.edo_id or "").strip() or None,
            "okved": (payload.okved or "").strip() or None,
        }


def _convert_office_to_pdf(source_bytes: bytes, suffix: str) -> bytes:
    """Один вызов soffice для docx договора и xlsx спецификации."""
    soffice = shutil.which("soffice")
    if not soffice:
        raise SupplyContractPdfError(MSG_LIBREOFFICE)
    with tempfile.TemporaryDirectory() as tmp:
        folder = Path(tmp)
        source = folder / f"document{suffix}"
        source.write_bytes(source_bytes)
        profile = folder / "profile"
        profile.mkdir()
        try:
            completed = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--norestore",
                    "--nolockcheck",
                    f"-env:UserInstallation={profile.as_uri()}",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(folder),
                    str(source),
                ],
                timeout=_PDF_TIMEOUT_SECONDS,
                capture_output=True,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            _log.error("soffice не собрал pdf: %s", exc)
            raise SupplyContractPdfError(MSG_LIBREOFFICE) from exc
        pdf_path = folder / "document.pdf"
        if completed.returncode != 0 or not pdf_path.is_file():
            _log.error(
                "soffice завершился с кодом %s: %s",
                completed.returncode,
                completed.stderr.decode("utf-8", errors="replace")[:500],
            )
            raise SupplyContractPdfError(MSG_LIBREOFFICE)
        return pdf_path.read_bytes()


def convert_docx_to_pdf(docx_bytes: bytes) -> bytes:
    """Конвертирует только что собранный docx. Это не GsmExportService."""
    return _convert_office_to_pdf(docx_bytes, ".docx")


def convert_xlsx_to_pdf(xlsx_bytes: bytes) -> bytes:
    """Конвертирует только что собранную книгу спецификации."""
    return _convert_office_to_pdf(xlsx_bytes, ".xlsx")


def _as_download(
    payload: bytes,
    number: str,
    stem: str,
    file_format: str,
) -> tuple[bytes, str]:
    if file_format not in _DOWNLOAD_FORMATS:
        raise SupplyContractValidationError("Укажите формат docx или pdf")
    filename = f"{stem}-{number.replace('/', '-')}.{file_format}"
    if file_format == "pdf":
        payload = convert_docx_to_pdf(payload)
    return payload, filename


def _require_edo_buyer(contract: dict[str, Any]) -> None:
    for key in _EDO_REQUIRED:
        if not str(contract.get(key) or "").strip():
            raise SupplyContractValidationError(MSG_EDO_REQUISITES)


def already_exists_message(number: str) -> str:
    return duplicate_contract_message(number)


def _imported_view(row: dict[str, Any], parties: list[dict[str, Any]]) -> dict[str, Any]:
    data = _present(row)
    data["imported_name"] = row.get("imported_name") or ""
    data["suggestions"] = suggest_names(str(data["imported_name"]), parties)
    data["counterparty_id"] = row.get("counterparty_id")
    return data


def _present(row: dict[str, Any] | None) -> dict[str, Any]:
    if row is None:
        return {}
    data = dict(row)
    path = data.pop("scan_path", None)
    data["has_scan"] = bool(path)
    return data


def _parse_date(value: date | str) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(value)


def _scan_suffix(data: bytes, filename: str) -> str:
    suffix = Path(filename or "").suffix.lower()
    if data.startswith(b"%PDF"):
        return ".pdf"
    if data.startswith(b"\xff\xd8\xff"):
        return ".jpeg" if suffix == ".jpeg" else ".jpg"
    if len(data) >= 8 and data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    raise SupplyContractValidationError("Неподдерживаемый формат скана. Допустимы PDF, JPEG или PNG.")


def _delete_owned_scan(scans_dir: Path, stored: str) -> None:
    path = Path(stored)
    try:
        path.resolve().relative_to(scans_dir.resolve())
    except ValueError:
        return
    path.unlink(missing_ok=True)


def _require_counterparty_card(parsed: dict[str, Any]) -> None:
    fields = parsed.get("fields") or {}
    if any(fields.get(key) for key in _CARD_REQUISITES):
        return
    raise SupplyContractValidationError(MSG_NOT_A_CARD)


def _is_image(data: bytes) -> bool:
    if data.startswith(b"\xff\xd8\xff"):
        return True
    return len(data) >= 8 and data.startswith(b"\x89PNG\r\n\x1a\n")


def _default_text_provider() -> Any:
    """Текстовый чат того же OCR_PROVIDER, без картинки."""
    from core.ocr.pipeline import create_ocr_provider

    vision, _name, _label = create_ocr_provider()

    class _TextProvider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict[str, Any], float]:
            return await vision.complete_text_json(user_text=user_text)

    return _TextProvider()


def _default_card_provider() -> Any:
    """Провайдер зрения карточки — тот же OCR_PROVIDER, что у КП."""
    from core.ocr.pipeline import create_ocr_provider
    from core.supply_contract_card_ocr import fields_from_model

    vision, _name, _label = create_ocr_provider()

    class _OcrCardProvider:
        async def extract_contract_fields(
            self,
            *,
            user_text: str,
            image_base64: str,
            mime_type: str,
        ) -> tuple[dict[str, Any], float]:
            raw, cost = await vision.complete_vision_json(
                user_text=user_text,
                image_base64=image_base64,
                mime_type=mime_type,
            )
            return fields_from_model(raw), cost

        async def reread_contract_fields(
            self,
            *,
            user_text: str,
            image_base64: str,
            mime_type: str,
        ) -> tuple[dict[str, Any], float]:
            raw, cost = await vision.complete_vision_json(
                user_text=user_text,
                image_base64=image_base64,
                mime_type=mime_type,
            )
            return fields_from_model(raw), cost

    return _OcrCardProvider()
