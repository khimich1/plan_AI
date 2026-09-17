/** Типы API POST /api/v1/nomenclature/import-1c — синхронны с app/schemas/nomenclature.py. */

export const NOMENCLATURE_PRODUCT_KINDS = [
  "pile",
  "bridge_pile",
  "fbs",
  "stair_flight",
  "stair_step",
] as const;

export type NomenclatureProductKind = (typeof NOMENCLATURE_PRODUCT_KINDS)[number];

export const PRODUCT_KIND_LABELS: Record<NomenclatureProductKind, string> = {
  pile: "Сваи",
  bridge_pile: "Мостовые сваи",
  fbs: "ФБС",
  stair_flight: "Лестничные марши",
  stair_step: "Лестничные ступени",
};

export type GuidWriteOut = {
  product_kind: string;
  mark: string;
  field: string;
  guid: string;
};

export type AmbiguousMatchOut = {
  product_kind: string;
  mark: string | null;
  name: string;
  guids: string[];
  note: string;
};

export type DisappearedMarkOut = {
  product_kind: string;
  mark: string;
  guid_1c: string | null;
  match_status: string;
};

export type Unmatched1COut = {
  name: string;
  guid: string;
  row_index: number;
  source_file: string;
};

export type Import1cResponse = {
  product_kind: string;
  summary: string;
  new_guids_count: number;
  waiting_price: number;
  ambiguous_count: number;
  disappeared_count: number;
  unmatched_1c_count: number;
  unchanged: number;
  updated_guids_count: number;
  mode?: string;
  weights_updated?: number;
  list_limit: number;
  new_guids: GuidWriteOut[];
  updated_guids: GuidWriteOut[];
  ambiguous: AmbiguousMatchOut[];
  disappeared: DisappearedMarkOut[];
  unmatched_1c: Unmatched1COut[];
};

export type Create1cTask = {
  product_kind: string;
  mark: string;
  hint: string;
  field?: string;
};

export type PriceTask = {
  guid: string;
  product_kind: string;
  mark: string;
  name: string;
};

export type DuplicateCandidate = {
  guid: string;
  name: string;
  price: number | null;
};

export type DuplicateTask = {
  scope: string;
  key: string;
  product_kind: string;
  candidates: DuplicateCandidate[];
};

export type GuidTasksResponse = {
  to_create_1c: Create1cTask[];
  to_price: PriceTask[];
  duplicates: DuplicateTask[];
};

export type PriceQueueResolveResponse = {
  guid: string;
  mark: string;
  product_kind: string;
  price: number;
  state: string;
};

export type DuplicateResolveResponse = {
  scope: string;
  key: string;
  chosen_guid: string;
  match_status: string | null;
};

export const productKindLabel = (kind: string): string => {
  if (kind in PRODUCT_KIND_LABELS) {
    return PRODUCT_KIND_LABELS[kind as NomenclatureProductKind];
  }
  if (kind === "plate" || kind === "plates") {
    return "Плиты";
  }
  return kind;
};
