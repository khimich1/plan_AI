/** Типы API «График поставки» — синхронны с `app/schemas/delivery_schedule.py`. */

export type ScheduleStatus = "draft" | "active" | "completed";
export type TrafficLightStatus = "green" | "yellow" | "red";
export type DeliveryScheduleDocumentFmt = "xlsx" | "pdf";

export type BatchItemIn = {
  plate_id: number;
  qty: number;
};

export type BatchItemOut = {
  plate_id: number;
  qty: number;
  plate_name: string | null;
  /** qty партии > текущего qty позиции КП. */
  changed: boolean;
};

export type BatchIn = {
  name: string;
  deliver_from: string;
  deliver_to: string;
  produce_by: string;
  items: BatchItemIn[];
  sort_order?: number;
};

export type BatchOut = {
  id: number;
  name: string;
  deliver_from: string;
  deliver_to: string;
  produce_by: string;
  items: BatchItemOut[];
  sort_order: number;
  status: TrafficLightStatus | null;
  ready_date: string | null;
  hint: string | null;
  /** Хотя бы одна позиция партии `changed`. */
  changed: boolean;
};

export type DeliverySchedulePut = {
  invoice_number?: string | null;
  contract_number?: string | null;
  batches: BatchIn[];
};

export type DeliveryScheduleView = {
  id: number;
  kp_id: number;
  invoice_number: string | null;
  contract_number: string | null;
  status: ScheduleStatus;
  batches: BatchOut[];
  updated_at: string;
  /** True, если светофор недоступен — statuses у партий null. */
  traffic_light_degraded?: boolean;
};

export type BatchDraftItemOut = {
  plate_id: number;
  plate_name: string;
  qty: number;
};

export type BatchDraftOut = {
  name: string;
  deliver_from: string;
  deliver_to: string;
  produce_by: string;
  items: BatchDraftItemOut[];
};

export type UnmatchedRowOut = {
  row_number: number;
  reason: string;
  raw: Record<string, unknown> | null;
};

export type ImportDraftResponse = {
  batches: BatchDraftOut[];
  unmatched_rows: UnmatchedRowOut[];
};
