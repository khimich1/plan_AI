import { httpClient } from "@/shared/api/httpClient";
import type {
  CardCreatePayload,
  CardPatchPayload,
  DriverCreatePayload,
  DriverPatchPayload,
  GsmCard,
  GsmDriver,
  GsmRoute,
  GsmSettings,
  GsmStation,
  GsmVehicle,
  GsmWaybill,
  StationCreatePayload,
  StationPatchPayload,
  TransactionImportReport,
  VehicleCreatePayload,
  VehiclePatchPayload,
  WaybillCreatePayload,
  WaybillExportPayload,
  WaybillGeneratePayload,
  WaybillGenerateResult,
  WaybillListParams,
  WaybillPatchPayload,
} from "@/features/gsm/types/gsm";

const BASE = "/api/v1/gsm";

const JSON_HEADERS = { "Content-Type": "application/json" };

export const gsmApi = {
  listVehicles: (params?: { activeOnly?: boolean }) => {
    const search = new URLSearchParams();
    if (params?.activeOnly === false) search.set("active_only", "false");
    const qs = search.toString();
    return httpClient.get<GsmVehicle[]>(`${BASE}/vehicles${qs ? `?${qs}` : ""}`);
  },

  createVehicle: (payload: VehicleCreatePayload) =>
    httpClient.post<GsmVehicle>(`${BASE}/vehicles`, JSON.stringify(payload), JSON_HEADERS),

  patchVehicle: (id: number, payload: VehiclePatchPayload) =>
    httpClient.patch<GsmVehicle>(
      `${BASE}/vehicles/${id}`,
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  listDrivers: (params?: { activeOnly?: boolean }) => {
    const search = new URLSearchParams();
    if (params?.activeOnly === false) search.set("active_only", "false");
    const qs = search.toString();
    return httpClient.get<GsmDriver[]>(`${BASE}/drivers${qs ? `?${qs}` : ""}`);
  },

  createDriver: (payload: DriverCreatePayload) =>
    httpClient.post<GsmDriver>(`${BASE}/drivers`, JSON.stringify(payload), JSON_HEADERS),

  patchDriver: (id: number, payload: DriverPatchPayload) =>
    httpClient.patch<GsmDriver>(
      `${BASE}/drivers/${id}`,
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  listCards: (params?: { includeArchived?: boolean }) => {
    const search = new URLSearchParams();
    if (params?.includeArchived) search.set("include_archived", "true");
    const qs = search.toString();
    return httpClient.get<GsmCard[]>(`${BASE}/cards${qs ? `?${qs}` : ""}`);
  },

  createCard: (payload: CardCreatePayload) =>
    httpClient.post<GsmCard>(`${BASE}/cards`, JSON.stringify(payload), JSON_HEADERS),

  patchCard: (id: number, payload: CardPatchPayload) =>
    httpClient.patch<GsmCard>(`${BASE}/cards/${id}`, JSON.stringify(payload), JSON_HEADERS),

  listStations: () => httpClient.get<GsmStation[]>(`${BASE}/stations`),

  createStation: (payload: StationCreatePayload) =>
    httpClient.post<GsmStation>(`${BASE}/stations`, JSON.stringify(payload), JSON_HEADERS),

  patchStation: (id: number, payload: StationPatchPayload) =>
    httpClient.patch<GsmStation>(
      `${BASE}/stations/${id}`,
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  getSettings: () => httpClient.get<GsmSettings>(`${BASE}/settings`),

  putSettings: (payload: GsmSettings) =>
    httpClient.put<GsmSettings>(`${BASE}/settings`, JSON.stringify(payload), JSON_HEADERS),

  importTransactions: (files: File[]) => {
    const form = new FormData();
    for (const file of files) {
      form.append("files", file);
    }
    return httpClient.post<TransactionImportReport>(`${BASE}/transactions/import`, form);
  },

  listWaybills: (params: WaybillListParams) => {
    const search = new URLSearchParams({
      vehicle_id: String(params.vehicleId),
      from: params.periodFrom,
      to: params.periodTo,
    });
    return httpClient.get<GsmWaybill[]>(`${BASE}/waybills?${search.toString()}`);
  },

  generateWaybills: (payload: WaybillGeneratePayload) =>
    httpClient.post<WaybillGenerateResult>(
      `${BASE}/waybills/generate`,
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  listRoutes: (vehicleId: number) => {
    const search = new URLSearchParams({ vehicle_id: String(vehicleId) });
    return httpClient.get<GsmRoute[]>(`${BASE}/routes?${search.toString()}`);
  },

  patchWaybill: (id: number, payload: WaybillPatchPayload) =>
    httpClient.patch<GsmWaybill>(
      `${BASE}/waybills/${id}`,
      JSON.stringify(payload),
      JSON_HEADERS,
    ),

  createWaybill: (payload: WaybillCreatePayload) =>
    httpClient.post<GsmWaybill>(`${BASE}/waybills`, JSON.stringify(payload), JSON_HEADERS),

  exportWaybills: (payload: WaybillExportPayload) =>
    httpClient.download(
      `${BASE}/waybills/export`,
      `gsm_waybills_${payload.from}_${payload.to}.zip`,
      {
        method: "POST",
        body: JSON.stringify(payload),
        headers: JSON_HEADERS,
      },
    ),
};
