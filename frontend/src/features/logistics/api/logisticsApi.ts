import { httpClient } from "@/shared/api/httpClient";
import type {
  Carrier,
  CarrierMergeResponse,
  CreateShipmentPayload,
  LogisticsKpSearchResponse,
  PileCatalogEntry,
  ProposeResponse,
  ShipmentDetails,
  ShipmentFilters,
  ShipmentItemInput,
  ShipmentMutationResult,
  ShipmentRegistryRow,
  UpdateShipmentPayload,
  VehicleClass,
} from "@/features/logistics/types/logistics";

const BASE = "/api/v1/logistics";

const JSON_HEADERS = { "Content-Type": "application/json" };

const buildShipmentsQuery = (filters: ShipmentFilters): string => {
  const params = new URLSearchParams();
  if (filters.date_from) params.set("date_from", filters.date_from);
  if (filters.date_to) params.set("date_to", filters.date_to);
  if (filters.kp_id != null) params.set("kp_id", String(filters.kp_id));
  if (filters.carrier_id != null) params.set("carrier_id", String(filters.carrier_id));
  if (filters.delivery_type) params.set("delivery_type", filters.delivery_type);
  if (filters.status) params.set("status", filters.status);
  if (filters.no_upd) params.set("no_upd", "1");
  if (filters.attention) params.set("attention", "1");
  const qs = params.toString();
  return qs ? `?${qs}` : "";
};

export const logisticsApi = {
  listShipments: (filters: ShipmentFilters = {}) =>
    httpClient
      .get<{ items: ShipmentRegistryRow[]; count: number }>(
        `${BASE}/shipments${buildShipmentsQuery(filters)}`,
      )
      .then((response) => response.items),

  createShipment: (payload: CreateShipmentPayload) =>
    httpClient.post<ShipmentDetails>(`${BASE}/shipments`, JSON.stringify(payload), JSON_HEADERS),

  reuseTransport: (sourceId: number, payload: CreateShipmentPayload) =>
    httpClient.post<ShipmentDetails>(
      `${BASE}/shipments/${sourceId}/reuse-transport`,
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  getShipment: (id: number) => httpClient.get<ShipmentDetails>(`${BASE}/shipments/${id}`),

  updateShipment: (id: number, payload: UpdateShipmentPayload) =>
    httpClient.patch<ShipmentDetails>(
      `${BASE}/shipments/${id}`,
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  // Backend читает vehicle_class из query-string (app/api/v1/endpoints/logistics.py).
  proposeItems: (id: number, vehicleClass?: VehicleClass | null) =>
    httpClient.post<ProposeResponse>(
      `${BASE}/shipments/${id}/propose${vehicleClass ? `?vehicle_class=${vehicleClass}` : ""}`,
    ),

  confirmItems: (id: number, items: ShipmentItemInput[]) =>
    httpClient.put<ShipmentDetails>(
      `${BASE}/shipments/${id}/items`,
      JSON.stringify({ items }),
      JSON_HEADERS,
    ),

  completeShipment: (id: number) =>
    httpClient.post<ShipmentMutationResult>(`${BASE}/shipments/${id}/complete`, JSON.stringify({}), JSON_HEADERS),

  cancelShipment: (id: number) =>
    httpClient.post<ShipmentMutationResult>(`${BASE}/shipments/${id}/cancel`, JSON.stringify({}), JSON_HEADERS),

  downloadSheet: (id: number) =>
    httpClient.download(`${BASE}/shipments/${id}/sheet.xlsx`, `Лист_отгрузки_${id}.xlsx`),

  listCarriers: (params?: { q?: string; activeOnly?: boolean }) => {
    const search = new URLSearchParams();
    if (params?.q) search.set("q", params.q);
    if (params?.activeOnly) search.set("active", "1");
    const qs = search.toString();
    return httpClient
      .get<{ items: Carrier[]; count: number }>(`${BASE}/carriers${qs ? `?${qs}` : ""}`)
      .then((response) => response.items);
  },

  mergeCarrier: (id: number, intoId: number) =>
    httpClient.post<CarrierMergeResponse>(
      `${BASE}/carriers/${id}/merge`,
      JSON.stringify({ into_id: intoId }),
      JSON_HEADERS,
    ),

  searchPileCatalog: (q = "") => {
    const qs = q ? `?q=${encodeURIComponent(q)}` : "";
    return httpClient
      .get<{ items: PileCatalogEntry[]; count: number }>(`${BASE}/pile-catalog${qs}`)
      .then((response) => response.items);
  },

  searchKp: (params: { kpId?: number; customer?: string }) => {
    const search = new URLSearchParams();
    if (params.kpId != null) search.set("kp_id", String(params.kpId));
    if (params.customer) search.set("customer", params.customer);
    return httpClient.get<LogisticsKpSearchResponse>(`${BASE}/kp-search?${search.toString()}`);
  },
};
