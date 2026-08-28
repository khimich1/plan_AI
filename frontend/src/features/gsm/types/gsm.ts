/** Types aligned with app/schemas/gsm.py (Task 5 registry API). */

export type GsmTab = "overview" | "transactions" | "registries";

export type GsmVehicle = {
  id: number;
  name: string;
  plate_number: string;
  tank_volume_liters: number;
  norm_summer: number;
  norm_winter: number;
  primary_driver_id: number | null;
  is_active: boolean;
};

export type GsmDriver = {
  id: number;
  full_name: string;
  license_number: string;
  license_issued_at: string | null;
  personnel_number: string | null;
  snils: string | null;
  is_active: boolean;
};

export type GsmCard = {
  id: number;
  card_number: string;
  vehicle_id: number | null;
  assigned_at: string;
  archived_at: string | null;
};

export type GsmStation = {
  id: number;
  address: string;
  brand: string | null;
  lat: number | null;
  lon: number | null;
  geocode_source: string | null;
};

export type GsmSeasonMode = "summer" | "winter";

export type GsmSettings = {
  winter_start: string;
  hook_threshold_km: number;
  season_mode: GsmSeasonMode;
  season_switched_at: string | null;
};

export type VehicleCreatePayload = {
  name: string;
  plate_number: string;
  tank_volume_liters: number;
  norm_summer: number;
  norm_winter: number;
  primary_driver_id?: number | null;
  is_active?: boolean;
};

export type VehiclePatchPayload = Partial<VehicleCreatePayload>;

export type DriverCreatePayload = {
  full_name: string;
  license_number: string;
  license_issued_at?: string | null;
  personnel_number?: string | null;
  snils?: string | null;
  is_active?: boolean;
};

export type DriverPatchPayload = Partial<DriverCreatePayload>;

export type CardCreatePayload = {
  card_number: string;
  vehicle_id?: number | null;
  assigned_at?: string | null;
};

export type CardPatchPayload = {
  vehicle_id?: number | null;
  archive?: boolean | null;
};

export type StationCreatePayload = {
  address: string;
  brand?: string | null;
  lat?: number | null;
  lon?: number | null;
  geocode_source?: string | null;
};

export type StationPatchPayload = Partial<StationCreatePayload>;

export type FileImportReport = {
  filename: string;
  rows_total: number;
  rows_inserted: number;
  rows_duplicate: number;
  sum_liters: number;
  sum_amount: number;
  footer_liters: number | null;
  footer_amount: number | null;
  warnings: string[];
  unmatched_cards: string[];
};

export type TransactionImportReport = {
  files: FileImportReport[];
  rows_inserted: number;
  rows_duplicate: number;
};

/** Day-level warning codes from core.gsm.generator / gsm_waybill.warnings_json. */
export type WaybillWarningCode =
  | "weekend_anchor"
  | "hook_above_threshold"
  | "unsolvable"
  | "manual_intervention"
  | "balance_route"
  | "borrowed_route"
  | string;

export type ProblematicDay = {
  date: string;
  reason: string;
  detail: string;
  fuel_before: number;
  fuel_to_issue: number;
  tank_volume: number;
};

export type WaybillRouteLeg = {
  from?: string;
  to?: string;
  from_addr?: string;
  to_addr?: string;
  km: number;
  route_id?: number | null;
  station_id?: number | null;
  dep_time?: string | null;
  arr_time?: string | null;
};

export type WaybillWarningDetail = {
  code: string;
  detail: string;
};

export type GsmWaybill = {
  id: number;
  vehicle_id: number;
  date: string;
  driver_id: number;
  status: string;
  source: string;
  odometer_start: number | null;
  odometer_end: number | null;
  fuel_start: number | null;
  fuel_issued: number | null;
  fuel_end: number | null;
  km: number;
  route: WaybillRouteLeg[];
  warnings: WaybillWarningCode[];
  warning_details?: WaybillWarningDetail[];
  /** Set on PATCH response: how many following draft days were recalculated. */
  rechained_draft_days?: number;
};

export type WaybillGeneratePayload = {
  vehicle_id: number;
  period_from: string;
  period_to: string;
  force?: boolean;
  fuel_start?: number | null;
  odometer_start?: number | null;
};

export type WaybillGenerateResult = {
  waybills: GsmWaybill[];
  warnings: WaybillWarningCode[];
  days_created: number;
  problematic_days: ProblematicDay[];
  manual_days: number;
};

export type WaybillListParams = {
  vehicleId?: number;
  periodFrom: string;
  periodTo: string;
};

export type GsmTransaction = {
  ts: string;
  card_number: string;
  vehicle_id: number | null;
  service_type: string;
  fuel_grade: string | null;
  qty_liters: number | null;
  amount: number;
  station_id: number | null;
  address: string | null;
};

export type TransactionListResponse = {
  rows: GsmTransaction[];
  total_count: number;
  sum_liters: number;
  sum_amount: number;
};

export type TransactionListParams = {
  periodFrom: string;
  periodTo: string;
  vehicleId?: number;
  serviceType?: string;
};

export type VehiclePeriodStatus =
  | "no_data"
  | "needs_generation"
  | "has_red_days"
  | "drafts_pending"
  | "pending_export"
  | "ready";

export type FleetOverviewVehicle = {
  id: number;
  name: string;
  plate_number: string;
};

export type FleetOverviewRow = {
  vehicle: FleetOverviewVehicle;
  tx_count: number;
  tx_liters: number;
  tx_amount: number;
  tx_last_date: string | null;
  wb_count: number;
  wb_km: number;
  wb_fuel_issued: number;
  wb_last_date: string | null;
  red_days: number;
  draft_count: number;
  confirmed_count: number;
  exported_count: number;
  fuel_end_last: number | null;
  liters_diff: number;
  open_before: number;
  open_before_month: string | null;
  chain_broken: boolean;
  status: VehiclePeriodStatus;
};

export type OverviewParams = {
  periodFrom: string;
  periodTo: string;
};

export type WaybillBulkGeneratePayload = {
  vehicle_ids: number[];
  period_from: string;
  period_to: string;
  force?: boolean;
};

export type BulkGenerateVehicleError = {
  code: string;
  message: string;
};

export type BulkGenerateVehicleResult = {
  vehicle_id: number;
  ok: boolean;
  result?: WaybillGenerateResult | null;
  error?: BulkGenerateVehicleError | null;
};

export type BulkGenerateResult = {
  results: BulkGenerateVehicleResult[];
};

/** POST /gsm/waybills/export — aliases `from`/`to` match backend schema. */
export type WaybillExportPayload = {
  vehicle_ids: number[];
  from: string;
  to: string;
};

/** POST /gsm/report/usage — null vehicle_ids = all active vehicles. */
export type UsageReportPayload = {
  period_from: string;
  period_to: string;
  vehicle_ids: number[] | null;
};

export type GsmRoute = {
  id: number;
  vehicle_id: number;
  addr_a: string;
  addr_b: string;
  km: number;
  frequency: number;
  typical_station_ids: number[];
};

export type WaybillPatchPayload = {
  driver_id?: number | null;
  km?: number | null;
  route?: WaybillRouteLeg[] | null;
};

export type WaybillCreatePayload = {
  vehicle_id: number;
  date: string;
  driver_id: number;
  route: WaybillRouteLeg[];
  fuel_issued?: number;
  fuel_start?: number | null;
  odometer_start?: number | null;
};
