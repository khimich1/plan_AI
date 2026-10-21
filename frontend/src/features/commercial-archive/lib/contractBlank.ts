import type {
  ContractCard,
  SupplyContractBankBundle,
  SupplyContractCreatePayload,
  SupplyContractFormValues,
  SupplyContractParseResult,
} from "@/features/commercial-archive/types/supplyContract";

const LEGAL_FORMS = ["ooo", "ao", "ip", "kfh", "person"] as const;

const LIST_VALUES: Record<string, readonly string[]> = {
  legal_form: LEGAL_FORMS,
  signatory_verb: ["действующего", "действующей"],
  authority_basis: ["устав", "доверенность"],
};

const REVIEW_LABELS = [
  ["inn", "ИНН"],
  ["kpp", "КПП"],
  ["ogrn", "ОГРН"],
  ["account", "расчётный счёт"],
  ["corr_account", "корсчёт"],
  ["bik", "БИК"],
  ["legal_address", "юридический адрес"],
  ["bank_name", "банк"],
  ["signatory_position", "должность"],
  ["signatory_name", "ФИО"],
] as const;

const BANK_PHRASE_KEYS = new Set(["bank_name", "account", "corr_account", "bik"]);

const isLegalForm = (
  value: string | null,
): value is SupplyContractCreatePayload["legal_form"] =>
  LEGAL_FORMS.includes(value as (typeof LEGAL_FORMS)[number]);

const acceptedField = (key: string, value: unknown) => {
  if (key === "contract_date" || value == null || value === "") {
    return false;
  }
  const allowed = LIST_VALUES[key];
  if (!allowed) {
    return true;
  }
  return allowed.includes(String(value));
};

const textOrEmpty = (value: string | null | undefined) => value ?? "";

export const blankFormFromCard = (source: ContractCard): SupplyContractFormValues => ({
  legal_form: isLegalForm(source.legal_form) ? source.legal_form : "",
  full_name: textOrEmpty(source.full_name),
  short_name: textOrEmpty(source.short_name),
  signatory_position: textOrEmpty(source.signatory_position),
  signatory_name: textOrEmpty(source.signatory_name),
  signatory_verb: source.signatory_verb === "действующей" ? "действующей" : "действующего",
  authority_basis: source.authority_basis === "доверенность" ? "доверенность" : "устав",
  poa_number: textOrEmpty(source.poa_number),
  poa_date: textOrEmpty(source.poa_date),
  inn: textOrEmpty(source.inn),
  kpp: textOrEmpty(source.kpp),
  ogrn: textOrEmpty(source.ogrn),
  legal_address: textOrEmpty(source.legal_address),
  postal_address: textOrEmpty(source.postal_address),
  phone: textOrEmpty(source.phone),
  email: textOrEmpty(source.email),
  bank_name: textOrEmpty(source.bank_name),
  account: textOrEmpty(source.account),
  corr_account: textOrEmpty(source.corr_account),
  bik: textOrEmpty(source.bik),
  edo_operator: textOrEmpty(source.edo_operator),
  edo_id: textOrEmpty(source.edo_id),
  okved: textOrEmpty(source.okved),
});

export const requisitesPayload = (
  form: SupplyContractFormValues,
): SupplyContractCreatePayload | null => {
  if (form.legal_form === "") {
    return null;
  }
  return {
    legal_form: form.legal_form,
    full_name: form.full_name,
    short_name: form.short_name,
    signatory_position: form.signatory_position || null,
    signatory_name: form.signatory_name,
    signatory_verb: form.signatory_verb,
    authority_basis: form.authority_basis,
    poa_number: form.poa_number || null,
    poa_date: form.poa_date || null,
    inn: form.inn,
    kpp: form.kpp || null,
    ogrn: form.ogrn,
    legal_address: form.legal_address,
    postal_address: form.postal_address || null,
    phone: form.phone || null,
    email: form.email,
    bank_name: form.bank_name,
    account: form.account,
    corr_account: form.corr_account,
    bik: form.bik,
    edo_operator: form.edo_operator || null,
    edo_id: form.edo_id || null,
    okved: form.okved?.trim() ? form.okved.trim() : null,
  };
};

export const mergeParsedFields = (
  current: SupplyContractFormValues,
  parsed: SupplyContractParseResult,
): SupplyContractFormValues => {
  const accounts = parsed.accounts ?? [];
  const banks = parsed.banks ?? [];
  const next: SupplyContractFormValues = {
    ...current,
    ...(Object.fromEntries(
      Object.entries(parsed.fields).filter(([key, value]) => acceptedField(key, value)),
    ) as Partial<SupplyContractFormValues>),
  };
  if (accounts.length > 1) {
    next.account = "";
  }
  if (banks.length > 1) {
    next.bank_name = "";
    next.account = "";
    next.corr_account = "";
    next.bik = "";
  }
  return next;
};

export const formWithBank = (
  form: SupplyContractFormValues,
  bank: SupplyContractBankBundle,
): SupplyContractFormValues => ({
  ...form,
  bank_name: bank.bank_name ?? "",
  account: bank.account,
  corr_account: bank.corr_account ?? "",
  bik: bank.bik,
});

export const digitWarningText = (
  doubtful: Set<string>,
  accountCount: number,
  bankCount: number,
) => {
  const labels = REVIEW_LABELS.filter(([key]) => {
    if (key === "account" && accountCount > 1) {
      return false;
    }
    if (bankCount > 1 && BANK_PHRASE_KEYS.has(key)) {
      return false;
    }
    return doubtful.has(key);
  }).map(([, label]) => label);
  if (labels.length === 0) {
    return null;
  }
  return `Проверьте по карточке: ${labels.join(", ")}.`;
};
