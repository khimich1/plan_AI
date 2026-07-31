import { httpClient } from "@/shared/api/httpClient";

const BASE = "/api/v1/production/sgp";

export type SgpFilter = "all" | "linked" | "unlinked";

export type SgpProgress = { n: number; m: number };

export type SgpPlateItem = {
  id: number;
  kp_id: number | null;
  plate_name: string;
  length_m: number | null;
  width_m: number | null;
  load_class: number | null;
  qty: number;
  completed_date: string | null;
  production_day: number | null;
  plan_id: string | null;
  nomenclature_id: string | null;
  customer_name: string | null;
  execution_terms: string | null;
  sgp_progress: SgpProgress | null;
};

export type SgpPlatesResponse = {
  items: SgpPlateItem[];
  count: number;
  filter: SgpFilter;
};

export type SgpMutationResponse = {
  ok: boolean;
  sgp_id: number;
  qty: number;
  kp_id: number | null;
  target_kp_id: number | null;
  message: string;
};

export type SgpFreePlateItem = {
  id: number;
  plate_name: string;
  length_m: number | null;
  width_m: number | null;
  load_class: number | null;
  qty: number;
  completed_date: string | null;
};

export type SgpFreePlatesResponse = {
  items: SgpFreePlateItem[];
  count: number;
};

export const sgpApi = {
  listPlates: (filter: SgpFilter = "all") =>
    httpClient.get<SgpPlatesResponse>(
      `${BASE}/plates?filter=${encodeURIComponent(filter)}`,
    ),

  freePlates: (params?: {
    plate_name?: string;
    length_m?: number;
    width_m?: number;
    load_class?: number;
  }) => {
    const q = new URLSearchParams();
    if (params?.plate_name) q.set("plate_name", params.plate_name);
    if (params?.length_m != null) q.set("length_m", String(params.length_m));
    if (params?.width_m != null) q.set("width_m", String(params.width_m));
    if (params?.load_class != null) q.set("load_class", String(params.load_class));
    const qs = q.toString();
    return httpClient.get<SgpFreePlatesResponse>(
      `${BASE}/free-plates${qs ? `?${qs}` : ""}`,
    );
  },

  unlink: (sgpId: number, qty: number) =>
    httpClient.post<SgpMutationResponse>(
      `${BASE}/plates/${sgpId}/unlink`,
      JSON.stringify({ qty }),
      { "Content-Type": "application/json" },
    ),

  relink: (sgpId: number, targetKpId: number, qty: number) =>
    httpClient.post<SgpMutationResponse>(
      `${BASE}/plates/${sgpId}/relink`,
      JSON.stringify({ target_kp_id: targetKpId, qty }),
      { "Content-Type": "application/json" },
    ),
};
