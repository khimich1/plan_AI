export type ArchiveSection = "archived" | "in_production" | "completed";
export type ArchiveFileKind = "pdf" | "xlsx" | "schema";

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
};

export type ArchivePlateItem = {
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

export type ArchiveOfferFinance = {
  subtotal: number;
  vat_amount: number;
  total_amount: number;
  discount_percent: number;
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
  plates: ArchivePlateItem[];
  completion_percentage: number | null;
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

export type ProductionEstimate = {
  total_length_m: number;
  estimated_tracks: number;
  estimated_days: number;
};
