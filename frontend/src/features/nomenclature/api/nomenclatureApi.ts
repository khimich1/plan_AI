import { httpClient } from "@/shared/api/httpClient";
import type {
  DuplicateResolveResponse,
  GuidTasksResponse,
  Import1cResponse,
  NomenclatureProductKind,
  PriceQueueResolveResponse,
} from "@/features/nomenclature/types/nomenclature";

const BASE = "/api/v1/nomenclature";
const JSON_HEADERS = { "Content-Type": "application/json" };

export const nomenclatureApi = {
  import1c: (file: File, productKind?: NomenclatureProductKind | null) => {
    const formData = new FormData();
    formData.append("file", file);
    if (productKind) {
      formData.append("product_kind", productKind);
    }
    return httpClient.post<Import1cResponse>(`${BASE}/import-1c`, formData);
  },

  listTasks: () => httpClient.get<GuidTasksResponse>(`${BASE}/tasks`),

  listDuplicates: () => httpClient.get<GuidTasksResponse>(`${BASE}/duplicates`),

  resolvePrice: (guid: string, price: number) =>
    httpClient.post<PriceQueueResolveResponse>(
      `${BASE}/price-queue/resolve`,
      JSON.stringify({ guid, price }),
      JSON_HEADERS,
    ),

  resolveDuplicate: (payload: {
    scope: string;
    key: string;
    chosen_guid: string;
    note?: string;
    product_kind?: string | null;
  }) =>
    httpClient.post<DuplicateResolveResponse>(
      `${BASE}/duplicates/resolve`,
      JSON.stringify({
        scope: payload.scope,
        key: payload.key,
        chosen_guid: payload.chosen_guid,
        note: payload.note ?? "",
        product_kind: payload.product_kind ?? null,
      }),
      JSON_HEADERS,
    ),
};
