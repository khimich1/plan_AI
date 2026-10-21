export const SUPPLY_CONTRACT_STATUSES = [
  "нет",
  "подписан по ЭДО",
  "оригинал в бухгалтерии",
  "подписан с синей печатью оригинал",
  "отмена не будем работать",
] as const;

export const CANCELLED_CONTRACT_STATUS = "отмена не будем работать";

export const SUPPLY_CONTRACT_SCAN_NOTES = ["", "получен скан", "внесены правки"] as const;

export type SupplyContractStatus = (typeof SUPPLY_CONTRACT_STATUSES)[number];

export type SupplyContract = {
  id: number;
  counterparty_id: number | null;
  number: string;
  contract_date: string | null;
  manager_name: string | null;
  status: string;
  scan_note?: string | null;
  has_scan?: boolean;
  legal_form: string | null;
  full_name: string | null;
  short_name: string | null;
  signatory_name: string | null;
  email: string | null;
  okved?: string | null;
};

export type SupplyContractRegistryRow = {
  id: number;
  number: string;
  contract_date: string | null;
  counterparty_name: string;
  counterparty_id: number | null;
  manager_name: string | null;
  status: string;
  scan_note: string | null;
  has_scan: boolean;
};

export type ContractCard = {
  id: number;
  name: string;
  imported_name: string | null;
  counterparty_id: number | null;
  number: string;
  contract_date: string | null;
  manager_name: string | null;
  status: string;
  scan_note: string | null;
  has_scan: boolean;
  legal_form: string | null;
  full_name: string | null;
  short_name: string | null;
  signatory_position: string | null;
  signatory_name: string | null;
  signatory_verb: string | null;
  authority_basis: string | null;
  poa_number: string | null;
  poa_date: string | null;
  inn: string | null;
  kpp: string | null;
  ogrn: string | null;
  legal_address: string | null;
  postal_address: string | null;
  phone: string | null;
  email: string | null;
  bank_name: string | null;
  account: string | null;
  corr_account: string | null;
  bik: string | null;
  edo_operator: string | null;
  edo_id: string | null;
  okved: string | null;
  requisites_complete: boolean;
};

export type ContractCardList = {
  items: ContractCard[];
  count: number;
};

export type ContractCardCreatePayload = {
  name: string;
  inn?: string | null;
  number?: string | null;
};

export type ContractCardPatchPayload = {
  name?: string | null;
  inn?: string | null;
  number?: string | null;
};

export type SupplyContractCreatePayload = {
  legal_form: "ooo" | "ao" | "ip" | "kfh" | "person";
  full_name: string;
  short_name: string;
  signatory_position?: string | null;
  signatory_name: string;
  signatory_verb: "действующего" | "действующей";
  authority_basis: "устав" | "доверенность";
  poa_number?: string | null;
  poa_date?: string | null;
  inn: string;
  kpp?: string | null;
  ogrn: string;
  legal_address: string;
  postal_address?: string | null;
  phone?: string | null;
  email: string;
  bank_name: string;
  account: string;
  corr_account: string;
  bik: string;
  edo_operator?: string | null;
  edo_id?: string | null;
  okved?: string | null;
  contract_date?: string | null;
};

export type SupplyContractFormValues = Omit<SupplyContractCreatePayload, "legal_form"> & {
  legal_form: SupplyContractCreatePayload["legal_form"] | "";
};

export type SupplyContractParsedFields = Partial<SupplyContractCreatePayload>;

export type SupplyContractBankBundle = {
  bank_name?: string | null;
  account: string;
  corr_account?: string | null;
  bik: string;
};

export type SupplyContractParseResult = {
  fields: SupplyContractParsedFields;
  doubtful: string[];
  accounts: string[];
  banks?: SupplyContractBankBundle[];
  source_text?: string;
  verify_failed: boolean;
};

export type SupplyContractPatchPayload = {
  status?: string;
  scan_note?: string | null;
  contract_date?: string;
  file?: File;
};
