"""Поля договора поставки, которые подтверждает менеджер."""

from __future__ import annotations

from datetime import date
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

LegalForm = Literal["ooo", "ao", "ip", "kfh", "person"]
SignatoryVerb = Literal["действующего", "действующей"]
AuthorityBasis = Literal["устав", "доверенность"]
ContractStatus = Literal[
    "нет",
    "подписан по ЭДО",
    "оригинал в бухгалтерии",
    "подписан с синей печатью оригинал",
    "отмена не будем работать",
]
ScanNote = Literal["получен скан", "внесены правки"]


class SupplyContractCreate(BaseModel):
    model_config = ConfigDict(extra="forbid")

    legal_form: LegalForm
    full_name: str = Field(min_length=1)
    short_name: str = Field(min_length=1)
    signatory_position: str | None = None
    signatory_name: str = Field(min_length=1)
    signatory_verb: SignatoryVerb = "действующего"
    authority_basis: AuthorityBasis
    poa_number: str | None = None
    poa_date: str | None = None
    inn: str = Field(min_length=1)
    kpp: str | None = None
    ogrn: str = Field(min_length=1)
    legal_address: str = Field(min_length=1)
    postal_address: str | None = None
    phone: str | None = None
    email: str = Field(min_length=1)
    bank_name: str = Field(min_length=1)
    account: str = Field(min_length=1)
    corr_account: str = Field(min_length=1)
    bik: str = Field(min_length=1)
    edo_operator: str | None = None
    edo_id: str | None = None
    okved: str | None = None
    contract_date: date | None = None


class SupplyContractParsedFields(BaseModel):
    """Поля формы после разбора. Пустые остаются пустыми, номер не входит."""

    model_config = ConfigDict(extra="ignore")

    legal_form: str | None = None
    full_name: str | None = None
    short_name: str | None = None
    signatory_position: str | None = None
    signatory_name: str | None = None
    signatory_verb: str | None = None
    authority_basis: str | None = None
    poa_number: str | None = None
    poa_date: str | None = None
    inn: str | None = None
    kpp: str | None = None
    ogrn: str | None = None
    legal_address: str | None = None
    postal_address: str | None = None
    phone: str | None = None
    email: str | None = None
    bank_name: str | None = None
    account: str | None = None
    corr_account: str | None = None
    bik: str | None = None
    edo_operator: str | None = None
    edo_id: str | None = None
    okved: str | None = None


class SupplyContractBankOut(BaseModel):
    model_config = ConfigDict(extra="ignore")

    bank_name: str | None = None
    account: str
    corr_account: str | None = None
    bik: str


class SupplyContractParseOut(BaseModel):
    model_config = ConfigDict(extra="ignore")

    fields: SupplyContractParsedFields
    doubtful: list[str]
    accounts: list[str] = Field(default_factory=list)
    banks: list[SupplyContractBankOut] = Field(default_factory=list)
    source_text: str = ""
    verify_failed: bool = False


class SupplyContractPatch(BaseModel):
    """Смена статуса, отметки скана или даты. Номер не входит и не пересчитывается."""

    model_config = ConfigDict(extra="forbid")

    status: ContractStatus | None = None
    scan_note: ScanNote | Literal[""] | None = None
    contract_date: date | None = None


class SupplyContractOut(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: int
    counterparty_id: int | None = None
    number: str
    contract_date: str | None = None
    manager_name: str | None = None
    status: str
    scan_note: str | None = None
    has_scan: bool = False
    legal_form: str | None = None
    full_name: str | None = None
    short_name: str | None = None
    signatory_position: str | None = None
    signatory_name: str | None = None
    signatory_verb: str | None = None
    authority_basis: str | None = None
    poa_number: str | None = None
    poa_date: str | None = None
    inn: str | None = None
    kpp: str | None = None
    ogrn: str | None = None
    legal_address: str | None = None
    postal_address: str | None = None
    phone: str | None = None
    email: str | None = None
    bank_name: str | None = None
    account: str | None = None
    corr_account: str | None = None
    bik: str | None = None
    edo_operator: str | None = None
    edo_id: str | None = None
    okved: str | None = None


class SupplyContractImportSuggestion(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: int
    name: str


class SupplyContractImportRow(BaseModel):
    """Строка листа до связки. suggestions не заполняют counterparty_id."""

    model_config = ConfigDict(extra="ignore")

    id: int
    number: str
    contract_date: str
    imported_name: str
    manager_name: str
    status: str
    scan_note: str | None = None
    counterparty_id: int | None = None
    suggestions: list[SupplyContractImportSuggestion] = Field(default_factory=list)


class SupplyContractLinkRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    counterparty_id: int = Field(ge=1)


class SupplyContractRegistryItem(BaseModel):
    """Колонки листа юриста. Путь файла скана наружу не отдаётся."""

    model_config = ConfigDict(extra="ignore")

    id: int
    number: str
    contract_date: str | None = None
    counterparty_name: str
    counterparty_id: int | None = None
    manager_name: str | None = None
    status: str
    scan_note: str | None = None
    has_scan: bool = False


class ContractCardCreate(BaseModel):
    """Новая карточка: имя обязательно, ИНН и номер можно не передавать."""

    model_config = ConfigDict(extra="forbid")

    name: str
    inn: str | None = None
    number: str | None = None


class ContractCardPatch(BaseModel):
    """Частичная правка имени, ИНН и номера. Непереданное поле строку не меняет."""

    model_config = ConfigDict(extra="forbid")

    name: str | None = None
    inn: str | None = None
    number: str | None = None


class ContractCardOut(SupplyContractOut):
    """Карточка с пустыми реквизитами. Путь скана наружу не отдаётся."""

    name: str
    imported_name: str | None = None
    requisites_complete: bool

    @model_validator(mode="before")
    @classmethod
    def _fill_has_scan(cls, value: Any) -> Any:
        if not isinstance(value, dict) or "has_scan" in value:
            return value
        data = dict(value)
        data["has_scan"] = bool(data.get("scan_path"))
        return data


class ContractCardListOut(BaseModel):
    model_config = ConfigDict(extra="ignore")

    items: list[ContractCardOut]
    count: int


class ContractCardImportOut(BaseModel):
    """Сколько карточек создано и сколько номеров пропущено."""

    model_config = ConfigDict(extra="ignore")

    created: int = Field(ge=0)
    skipped: int = Field(ge=0)
