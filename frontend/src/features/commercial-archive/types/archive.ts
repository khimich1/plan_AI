export type ArchiveSection = "archived" | "in_production" | "completed";
export type ArchiveFileKind = "pdf" | "xlsx" | "schema" | "xlsx_delivery_in_unit";
export type ProductType =
  | "plates"
  | "piles"
  | "steps"
  | "marches"
  | "bridge_piles"
  | "composite_piles"
  | "fbs"
  | "mixed";
export type ArchiveProductTypeFilter =
  | "all"
  | "plates"
  | "piles"
  | "steps"
  | "marches"
  | "bridge_piles"
  | "composite_piles"
  | "fbs";

export type ArchiveOfferListItem = {
  kp_id: number;
  creation_date: string | null;
  customer_name: string | null;
  manager_name: string | null;
  discount_percent: number;
  subtotal: number;
  vat_amount: number;
  total_amount: number;
  execution_terms: string | null;
  status: string | null;
  completion_percentage: number | null;
  sgp_progress?: { n: number; m: number } | null;
  /** Отгружено рейсами «обработано»: x из m (m = ordered_qty КП). */
  shipped_progress?: { x: number; m: number } | null;
  product_type?: ProductType;
  /** Concrete types for multi badges (Q3); preferred over a single mixed product_type. */
  product_types?: ProductType[];
  counterparty_id?: number | null;
  /** Optional: backend may omit — list badge skipped when absent. */
  has_delivery_schedule?: boolean;
};

export type ArchivePlateItem = {
  /** kp_plates.id — нужен для графика поставки; может отсутствовать у агрегатов. */
  id?: number | null;
  position_number: number | null;
  plate_name: string;
  length_m: number | null;
  width_m: number | null;
  load_class: number | null;
  qty: number;
  unit_price: number | null;
  discounted_price: number | null;
  unit_weight: number | null;
  total_weight: number | null;
  status: string | null;
  frost_resistance?: string | null;
  waterproofness?: string | null;
  concrete_aggregate?: string | null;
  concrete_spec_source?: string | null;
};

export type ArchivePileItem = {
  position_number: number | null;
  mark: string;
  concrete_grade: string;
  qty: number;
  unit_price: number | null;
  discounted_price: number | null;
  frost_resistance?: string | null;
  waterproofness?: string | null;
  concrete_aggregate?: string | null;
  concrete_spec_source?: string | null;
};

export type ArchiveStepItem = {
  position_number: number | null;
  mark: string;
  qty: number;
  unit_price: number | null;
  discounted_price: number | null;
};

export type ArchiveMarchItem = {
  position_number: number | null;
  mark: string;
  concrete_grade: string;
  qty: number;
  unit_price: number | null;
  discounted_price: number | null;
  frost_resistance?: string | null;
  waterproofness?: string | null;
  concrete_aggregate?: string | null;
  concrete_spec_source?: string | null;
};

export type ArchiveBridgePileItem = {
  position_number: number | null;
  mark: string;
  concrete_grade: string;
  qty: number;
  unit_price: number | null;
  discounted_price: number | null;
  frost_resistance?: string | null;
  waterproofness?: string | null;
  concrete_aggregate?: string | null;
  concrete_spec_source?: string | null;
};

export type ArchiveOfferFinance = {
  subtotal: number;
  vat_amount: number;
  total_amount: number;
  discount_percent: number;
};

export type SgpProgress = {
  n: number;
  m: number;
};

export type KpReadinessStepState = "done" | "active" | "pending" | "disabled";

export type KpReadinessStep = {
  id: "kp" | "production" | "sgp" | "release" | "closed";
  label: string;
  state: KpReadinessStepState;
  hint?: string | null;
};

export type KpReadinessSummary = {
  completion_percentage: number | null;
  sgp_progress: SgpProgress | null;
  issuable_qty: number;
  in_production_qty: number;
  summary_text: string;
  client_copy_text: string;
  steps: KpReadinessStep[];
  release_note?: string | null;
  expected_sgp_date?: string | null;
  expected_sgp_date_label?: string | null;
  fully_scheduled?: boolean;
};

export type KpReadinessPositionItem = {
  position_number: number | null;
  plate_name: string;
  length_m: number | null;
  width_m: number | null;
  load_class: number | null;
  label: string;
  ordered: number;
  in_plan: number;
  on_sgp: number;
  remaining: number;
};

export type KpReadinessPositionsResponse = {
  items: KpReadinessPositionItem[];
  count: number;
};

export type ArchiveOfferDetails = {
  kp_id: number;
  creation_date: string | null;
  customer_name: string | null;
  customer_inn?: string | null;
  customer_kpp?: string | null;
  counterparty_id?: number | null;
  manager_name: string | null;
  status: string | null;
  execution_terms: string | null;
  delivery_conditions: string | null;
  payment_conditions: string | null;
  finance: ArchiveOfferFinance;
  /** Стоимость одного рейса — то же поле, что logistics_cost при создании КП. */
  logistics_cost: number;
  pile_logistics_cost?: number;
  pile_trip_overrides?: Record<string, number>;
  pile_trips?: number;
  pile_trip_pending_marks?: string[];
  pile_delivery_ready?: boolean;
  plate_delivery_total?: number;
  pile_delivery_total?: number;
  fbs_lm_delivery_total?: number;
  fbs_lm_cargo_kg?: number;
  fbs_lm_trips?: number;
  fbs_lm_delivery_ready?: boolean;
  fbs_lm_pending_marks?: string[];
  fbs_lm_delivery_enabled?: boolean;
  long_pile_delivery_enabled?: boolean;
  long_pile_delivery_total?: number;
  long_pile_lengths?: Array<{
    length_key: number;
    trips: number;
    trip_cost: number | null;
    pending_marks: string[];
    ready: boolean;
    amount: number;
    qty: number;
  }>;
  long_pile_pending_marks?: string[];
  /** Масса груза (кг) по тем же правилам, что PDF/XLSX (resolve_kp_line_weight_kg на бэкенде). */
  total_cargo_weight_kg: number;
  /** Строка «Услуга по доставке грузов» = logistics_cost × число рейсов. */
  delivery_service_total_rub: number;
  product_type?: ProductType;
  plates: ArchivePlateItem[];
  piles?: ArchivePileItem[];
  steps?: ArchiveStepItem[];
  marches?: ArchiveMarchItem[];
  bridge_piles?: ArchiveBridgePileItem[];
  composite_piles?: ArchiveBridgePileItem[];
  fbs?: ArchiveBridgePileItem[];
  completion_percentage: number | null;
  readiness?: KpReadinessSummary | null;
};

export type ArchiveSearchState =
  | { kind: "number"; value: number }
  | { kind: "customer"; value: string }
  | null;

export type ArchiveSearchResponse = {
  mode: "number" | "customer";
  items: ArchiveOfferListItem[];
  total: number;
  truncated: boolean;
};

/** Ответ /archive/search для admin/manager. */
export type ArchiveSearchApiResponse = ArchiveSearchResponse;

export type ProductionEstimate = {
  total_length_m: number;
  estimated_tracks: number;
  estimated_days: number;
};
