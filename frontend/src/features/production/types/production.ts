export interface PlanMetaSummary {
  id: string;
  name: string;
  created_at?: string;
  start_date?: string;
  tracks_count?: number;
  total_days?: number;
  total_tracks?: number;
  completed_days?: number[];
  /** Optimistic-lock version from SQLite production_plans. */
  version?: number;
  [key: string]: unknown;
}

export interface PlansMetadataResponse {
  plans: PlanMetaSummary[];
  active_plan_id: string | null;
}

export interface DayInfo {
  occupied: number;
  max: number;
  completed: boolean;
  day_number: number;
}

export interface GlobalCalendarResponse {
  start_date?: string;
  total_days?: number;
  days_info: Record<string, DayInfo>;
  completed_days: number[];
  plans_count: number;
  tracks_count?: number;
}

export interface DayPlateInfo {
  customer: string;
  plate_name: string;
  kp_date: string;
  kp_id: number | null;
  length_m: number;
  width_mm: number;
  qty: number;
  reinforcement: number;
  load_code: number | null;
  /** Позиция из снимка после списания (остаётся в списке дня). */
  write_off_completed?: boolean;
}

export interface DayTrackDetail {
  track_number: number;
  /** 0-based индекс дорожки внутри плана (DELETE API). */
  plan_track_index: number;
  length: number | null;
  max_reinforcement: number;
  label: string | null;
  source_plan_id: string | null;
  source_plan_name: string | null;
  plates_info: DayPlateInfo[];
}

export interface DayPlanBlock {
  plan_id: string;
  plan_name: string;
  completed: boolean;
  tracks: DayTrackDetail[];
}

export interface DayViewResponse {
  date: string;
  plans: DayPlanBlock[];
  plans_count: number;
  total_tracks: number;
}

export interface CompleteDayResponse {
  plan_id: string;
  date: string;
  completed: boolean;
  moved_plates?: number;
  rejected_returned?: number;
  planned_qty_total?: number;
  completed_requested_qty?: number;
  rejected_requested_qty?: number;
  completed_kps?: number[];
  affected_kps?: number[];
  day_number?: number;
  rejected_plates?: number;
  rejected_positions?: number;
  skipped_without_kp_count?: number;
}

export interface RejectedPlateItem {
  track_number: number;
  plate_index: number;
  qty: number;
}

export type DayDocumentKind = "schema" | "breakdown" | "formovka";

export interface DayOccupancyResponse {
  occupancy: Record<string, number>;
  max_per_day: number;
}

export interface KpCandidatePlateItem {
  id: number;
  plate_name: string;
  length_m: number;
  width_m: number;
  load_class: number | null;
  qty: number;
}

export interface KpCandidateItem {
  kp_id: number;
  customer_name: string;
  creation_date: string;
  execution_terms: string;
  total_plates: number;
  completed_plates: number;
  completion_pct: number;
  in_plan_pct: number;
  total_length_m: number;
  plates: KpCandidatePlateItem[];
}

export interface KpCandidatesResponse {
  items: KpCandidateItem[];
  count: number;
}

export type FilterMethod = "all" | "kp";

export interface FillTargetItem {
  date: string;
  tracks: number;
}

export type LayoutReinforcementOrder = "asc" | "desc";

export interface BuildPlanRequest {
  start_date: string;
  tracks_count: number;
  filter_method: FilterMethod;
  selected_kp_ids?: number[];
  selected_plate_ids?: Record<number, number[]>;
  /** КП → id строки kp_plates → количество для этого плана (не больше qty в БД). */
  selected_plate_qty?: Record<number, Record<number, number>>;
  active_plan_id?: string | null;
  /** Optimistic-lock version активного плана при дозаписи существующего плана. */
  expected_version?: number;
  plan_name?: string | null;
  fill_targets?: FillTargetItem[];
  /** asc — слабые первыми (по умолчанию); desc — сильные первыми (экспериментальный режим). */
  layout_reinforcement_order?: LayoutReinforcementOrder;
}

export interface BuildPlanSummary {
  total_tracks: number;
  total_days: number;
  selected_plates_count: number;
  kp_count: number;
}

export interface BuildPlanResponse {
  plan: Record<string, unknown> & { id?: string; name?: string; version?: number };
  stats: Record<string, unknown>;
  summary: BuildPlanSummary;
}

export interface DeletePlanResponse {
  plan_id: string;
  deleted: boolean;
}

/** GET /production/plans/{id} — payload плана + version. */
export type PlanDetailResponse = Record<string, unknown> & {
  id: string;
  version: number;
};

export interface RemoveTrackResponse {
  plan_id: string;
  date: string;
  track_index: number;
  plates_returned: number;
  saved_tracks_count: number;
  warnings?: string[] | null;
}

export interface WorkCalendarPayload {
  extra_holidays: string[];
  extra_workdays: string[];
}

export type ProductionTab = "calendar" | "create" | "plans" | "work-calendar";
