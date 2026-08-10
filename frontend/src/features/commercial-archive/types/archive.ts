export type ArchiveSection = "archived" | "in_production" | "completed";
export type ArchiveFileKind = "pdf" | "xlsx" | "schema";
export type ProductType = "plates" | "piles" | "steps" | "marches" | "bridge_piles" | "fbs";
export type ArchiveProductTypeFilter = "all" | "plates" | "piles" | "steps" | "marches" | "bridge_piles" | "fbs";

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
};

export type ArchivePileItem = {
  position_number: number | null;
  mark: string;
  concrete_grade: string;
  qty: number;
  unit_price: number | null;
  discounted_price: number | null;
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
};

export type ArchiveBridgePileItem = {
  position_number: number | null;
  mark: string;
  concrete_grade: string;
  qty: number;
  unit_price: number | null;
  discounted_price: number | null;
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
  manager_name: string | null;
  status: string | null;
  execution_terms: string | null;
  delivery_conditions: string | null;
  payment_conditions: string | null;
  finance: ArchiveOfferFinance;
  /** Стоимость одного рейса — то же поле, что logistics_cost при создании КП. */
  logistics_cost: number;
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
