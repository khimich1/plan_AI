/** Типы API стола прайсов — синхронны с app/schemas/price_desk.py. */

export const PRICE_DESK_KINDS = [
  "plates",
  "fbs",
  "march",
  "step",
  "bridge_pile",
  "pile",
] as const;

export type PriceDeskKind = (typeof PRICE_DESK_KINDS)[number];

export const PRICE_DESK_KIND_LABELS: Record<PriceDeskKind, string> = {
  plates: "Плиты ПБ",
  fbs: "ФБС",
  march: "ЛМ",
  step: "Ступени",
  bridge_pile: "Мостовые сваи",
  pile: "Цельные сваи",
};

export type PriceDeskDiffExample = {
  key: string;
  old_price: number | null;
  new_price: number | null;
};

export type PriceDeskDiffExamples = {
  changed: PriceDeskDiffExample[];
  new: PriceDeskDiffExample[];
  missing: PriceDeskDiffExample[];
  unchanged: PriceDeskDiffExample[];
};

export type PriceDeskPreview = {
  product_kind: PriceDeskKind;
  price_list_date: string | null;
  file_sha256: string;
  parsed_rows: number;
  changed: number;
  new: number;
  missing: number;
  unchanged: number;
  examples: PriceDeskDiffExamples;
};

export type PriceGroupStatus = {
  product_kind: PriceDeskKind;
  price_list_date: string | null;
  imported_at: string | null;
  row_count: number;
};

export type PriceDeskStatus = {
  groups: PriceGroupStatus[];
};

export const priceDeskKindLabel = (kind: string): string => {
  if (kind in PRICE_DESK_KIND_LABELS) {
    return PRICE_DESK_KIND_LABELS[kind as PriceDeskKind];
  }
  return kind;
};
