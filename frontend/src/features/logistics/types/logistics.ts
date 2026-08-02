export type DeliveryType = "delivery" | "pickup";
export type ShipmentStatus = "in_work" | "done";
export type ShipmentItemType = "plate" | "free";
export type VehicleClass = "t20" | "t30plus";
export type LogisticsProductType = "plates" | "piles";

/** Slim KP search item for logistics (no financial fields). */
export type LogisticsKpSearchItem = {
  kp_id: number;
  customer_name: string | null;
  status: string | null;
  product_type?: LogisticsProductType;
};

export type LogisticsKpSearchResponse = {
  mode: "number" | "customer";
  items: LogisticsKpSearchItem[];
  total: number;
  truncated: boolean;
};

/** SQLite-хранимый флаг: API может вернуть 0/1 или true/false. */
export type BoolFlag = boolean | number;

export type ShipmentOrderRef = {
  kp_id: number;
  ya_order_no: string | null;
  customer_name: string | null;
};

export type ShipmentRegistryRow = {
  id: number;
  shipment_date: string;
  delivery_type: DeliveryType;
  status: ShipmentStatus;
  attention: BoolFlag;
  attention_comment: string | null;
  carrier_name: string | null;
  proxy_no: string | null;
  driver_name: string | null;
  vehicle_text: string | null;
  upd_no: string | null;
  planned_cost: number | null;
  total_weight_kg: number | null;
  orders: ShipmentOrderRef[];
};

export type ShipmentOrder = ShipmentOrderRef & {
  id: number;
};

export type ShipmentItem = {
  id: number;
  item_type: ShipmentItemType;
  completed_plate_id: number | null;
  kp_id: number | null;
  mark: string | null;
  plate_name: string | null;
  length_m: number | null;
  width_m: number | null;
  load_class: number | null;
  qty: number;
  unit_weight_kg: number | null;
  weight_kg: number | null;
  sort_order: number;
  note: string | null;
};

export type AvailablePlate = {
  completed_plate_id: number;
  plate_name: string;
  length_m: number | null;
  width_m: number | null;
  load_class: number | null;
  available_qty: number;
  unit_weight_kg: number | null;
};

export type AvailableByKp = {
  kp_id: number;
  plates: AvailablePlate[];
};

export type ShipmentDetails = {
  id: number;
  shipment_date: string;
  delivery_type: DeliveryType;
  status: ShipmentStatus;
  attention: BoolFlag;
  attention_comment: string | null;
  carrier_id: number | null;
  carrier_name: string | null;
  driver_name: string | null;
  vehicle_text: string | null;
  vehicle_class: VehicleClass | null;
  proxy_no: string | null;
  upd_no: string | null;
  freight_request_no: string | null;
  planned_cost: number | null;
  total_weight_kg: number | null;
  completed_at: string | null;
  orders: ShipmentOrder[];
  items: ShipmentItem[];
  available_by_kp: AvailableByKp[];
};

/** Совпадает с backend ShipmentProposeItem (app/schemas/logistics.py). */
export type ProposedItem = {
  item_type: ShipmentItemType;
  completed_plate_id: number;
  kp_id: number | null;
  plate_name: string;
  length_m: number | null;
  width_m: number | null;
  load_class: number | null;
  qty: number;
  available_qty: number;
  unit_weight_kg: number | null;
  weight_kg: number | null;
  completed_date?: string | null;
  reason_code?: string | null;
  reason_text?: string | null;
};

export type ProposeWarning = {
  code: string;
  message: string;
  kp_ids?: number[];
};

export type OrderRemainderItem = {
  completed_plate_id: number;
  kp_id: number;
  plate_name: string;
  qty_remaining: number;
};

/** Совпадает с backend ShipmentLayout* (app/schemas/logistics.py). */
export type LayoutUnit = {
  completed_plate_id: number;
  kp_id: number;
  plate_name: string;
  width_m: number | null;
};

export type LayoutTier = {
  index: number;
  units: LayoutUnit[];
};

export type LayoutStack = {
  index: number;
  marking_length_m: number;
  tiers: LayoutTier[];
};

export type LoadingStep = {
  step: number;
  stack_index: number;
  tier_index: number;
  description: string;
};

export type LayoutMetadata = {
  body_length_m: number;
  body_used_m: number;
  stacks: LayoutStack[];
  loading_steps: LoadingStep[];
};

export type ProposeResponse = {
  items: ProposedItem[];
  not_fit: ProposedItem[];
  order_remainder?: OrderRemainderItem[];
  warnings?: ProposeWarning[];
  total_weight_kg: number;
  overload: boolean;
  vehicle_class?: VehicleClass | null;
  vehicle_class_limits_kg: Record<VehicleClass, number>;
  layout?: LayoutMetadata | null;
};

/** Ответ complete/cancel (backend ShipmentMutationResponse). */
export type ShipmentMutationResult = {
  ok: boolean;
  shipment_id: number;
  status: ShipmentStatus;
  message: string;
};

export type ShipmentItemInput = {
  item_type: ShipmentItemType;
  completed_plate_id?: number | null;
  kp_id?: number | null;
  mark?: string | null;
  qty: number;
  weight_kg?: number | null;
  sort_order: number;
  note?: string | null;
};

export type ShipmentOrderInput = {
  kp_id: number;
  ya_order_no?: string | null;
};

export type CreateShipmentPayload = {
  shipment_date: string;
  delivery_type: DeliveryType;
  kp_ids: number[];
};

export type UpdateShipmentPayload = Partial<{
  carrier_id: number | null;
  driver_name: string | null;
  vehicle_text: string | null;
  vehicle_class: VehicleClass | null;
  proxy_no: string | null;
  upd_no: string | null;
  freight_request_no: string | null;
  planned_cost: number | null;
  attention: boolean;
  attention_comment: string | null;
  shipment_date: string;
  delivery_type: DeliveryType;
  orders: ShipmentOrderInput[];
}>;

export type ShipmentFilters = {
  date_from?: string;
  date_to?: string;
  kp_id?: number;
  carrier_id?: number;
  delivery_type?: DeliveryType;
  status?: ShipmentStatus;
  no_upd?: boolean;
  attention?: boolean;
};

export type Carrier = {
  id: number;
  name: string;
  shipments_count: number;
  active: BoolFlag;
  source_sheet?: string | null;
  note?: string | null;
  merged_into_id?: number | null;
};

export type CarrierMergeResponse = {
  moved_shipments: number;
};

export type PileCatalogEntry = {
  id: number;
  mark: string;
  length_m: number | null;
  section_mm: number | null;
  volume_m3: number | null;
  weight_kg: number;
};
