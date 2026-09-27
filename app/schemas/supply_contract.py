"""Поля договора поставки, которые подтверждает менеджер."""

from __future__ import annotations

from datetime import date
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

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
    contract_date: str
    manager_name: str
    status: str
    scan_note: str | None = None
    has_scan: bool = False
    legal_form: str
    full_name: str
    short_name: str
    signatory_position: str | None = None
    signatory_name: str
    signatory_verb: str
    authority_basis: str
    poa_number: str | None = None
    poa_date: str | None = None
    inn: str
    kpp: str | None = None
    ogrn: str
    legal_address: str
    postal_address: str | None = None
    phone: str | None = None
    email: str
    bank_name: str
    account: str
    corr_account: str
    bik: str
    edo_operator: str | None = None
    edo_id: str | None = None


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
    contract_date: str
    counterparty_name: str
    counterparty_id: int | None = None
    manager_name: str
    status: str
    scan_note: str | None = None
    has_scan: bool = False
