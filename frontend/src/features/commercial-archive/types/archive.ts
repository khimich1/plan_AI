export type ArchiveSection = "archived" | "in_production" | "completed";
export type ArchiveFileKind = "pdf" | "xlsx";

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
  plates: ArchivePlateItem[];
  completion_percentage: number | null;
};

export type ArchiveSearchResponse = {
  found: boolean;
  offer: ArchiveOfferDetails | null;
};

export type ProductionEstimate = {
  total_length_m: number;
  estimated_tracks: number;
  estimated_days: number;
};
