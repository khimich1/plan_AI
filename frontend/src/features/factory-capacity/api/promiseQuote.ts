import { useMutation, useQueries, useQuery, useQueryClient } from "@tanstack/react-query";
import { httpClient } from "@/shared/api/httpClient";
import { ApiError } from "@/shared/lib/apiError";

const BASE = "/api/v1/commercial/archive";
const SETTINGS_BASE = "/api/v1/commercial/settings";
const JSON_HEADERS = { "Content-Type": "application/json" };

/** GET /archive/{kp_id}/promise-quote — mirrors PromiseQuoteResponse (T3). */
export type PromiseQuoteWindow = {
  from_week: string;
  to_week: string;
  promised_date: string;
};

export type PromiseQuoteWeek = {
  week_start: string;
  workdays: number;
  capacity: number;
  planned: number;
  promised: number;
  held: number;
  free: number;
};

export type PromiseQuote = {
  tracks: number;
  solo_days: number;
  solo_date: string | null;
  solo_week_end_date: string | null;
  earliest_start_week: string | null;
  window: PromiseQuoteWindow | null;
  weeks: PromiseQuoteWeek[];
  knob: number;
};

export const promiseQuoteApi = {
  get: (kpId: number) => httpClient.get<PromiseQuote>(`${BASE}/${kpId}/promise-quote`),
};

export const promiseQuoteKeys = {
  all: ["promise-quote"] as const,
  quote: (kpId: number) => ["promise-quote", kpId] as const,
};

export const usePromiseQuoteQuery = (kpId: number | null) =>
  useQuery<PromiseQuote>({
    queryKey: promiseQuoteKeys.quote(kpId ?? -1),
    queryFn: () => promiseQuoteApi.get(kpId as number),
    enabled: kpId !== null,
    staleTime: 10_000,
  });

/** POST/GET/DELETE /archive/{kp_id}/promise-hold — mirrors PromiseHoldResponse (T5). */
export type PromiseHoldAllocation = {
  week_start: string;
  tracks: number;
};

export type PromiseHold = {
  id: number;
  kp_id: number;
  kind: "hold";
  status: "active" | "consumed" | "released" | "expired";
  tracks_total: number;
  promised_date: string;
  expires_at: string;
  created_by: string | null;
  created_at: string;
  allocations: PromiseHoldAllocation[];
};

export const promiseHoldApi = {
  get: (kpId: number) => httpClient.get<PromiseHold>(`${BASE}/${kpId}/promise-hold`),
  create: (kpId: number) => httpClient.post<PromiseHold>(`${BASE}/${kpId}/promise-hold`),
  release: (kpId: number) => httpClient.delete<PromiseHold>(`${BASE}/${kpId}/promise-hold`),
};

export const promiseHoldKeys = {
  all: ["promise-hold"] as const,
  hold: (kpId: number) => ["promise-hold", kpId] as const,
};

export async function fetchPromiseHold(kpId: number): Promise<PromiseHold | null> {
  try {
    const hold = await promiseHoldApi.get(kpId);
    return hold.status === "active" ? hold : null;
  } catch (error) {
    if (error instanceof ApiError && error.status === 404) {
      return null;
    }
    throw error;
  }
}

export const usePromiseHoldQuery = (kpId: number | null) =>
  useQuery<PromiseHold | null>({
    queryKey: promiseHoldKeys.hold(kpId ?? -1),
    queryFn: () => fetchPromiseHold(kpId as number),
    enabled: kpId !== null,
    staleTime: 10_000,
    retry: false,
  });

/** Active holds for archive rows (GET per КП; 404 → нет холда). */
export const usePromiseHoldsMap = (kpIds: number[]): Map<number, PromiseHold> => {
  const unique = [...new Set(kpIds)];
  const results = useQueries({
    queries: unique.map((kpId) => ({
      queryKey: promiseHoldKeys.hold(kpId),
      queryFn: () => fetchPromiseHold(kpId),
      staleTime: 15_000,
      retry: false,
    })),
  });
  const map = new Map<number, PromiseHold>();
  unique.forEach((kpId, index) => {
    const hold = results[index]?.data;
    if (hold) {
      map.set(kpId, hold);
    }
  });
  return map;
};

const invalidateAfterHoldWrite = (queryClient: ReturnType<typeof useQueryClient>, kpId: number) => {
  queryClient.invalidateQueries({ queryKey: promiseQuoteKeys.all });
  queryClient.invalidateQueries({ queryKey: promiseHoldKeys.all });
  queryClient.invalidateQueries({ queryKey: promiseHoldKeys.hold(kpId) });
  queryClient.invalidateQueries({ queryKey: ["archive"] });
};

export const useCreatePromiseHoldMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (kpId: number) => promiseHoldApi.create(kpId),
    onSuccess: (hold) => {
      queryClient.setQueryData(promiseHoldKeys.hold(hold.kp_id), hold);
      invalidateAfterHoldWrite(queryClient, hold.kp_id);
    },
  });
};

/** Tooltip / title: who pinned the date (named visibility). */
export function holdCreatedByTitle(createdBy: string | null | undefined): string {
  const who = createdBy?.trim();
  return who ? `Закрепил: ${who}` : "Срок закреплён до сегодня";
}

/** ISO YYYY-MM-DD → «4.09» (спека: «обещать к 25.09»). */
export function formatQuoteDayMonth(iso: string | null | undefined): string {
  if (!iso) return "—";
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!match) return iso;
  return `${Number(match[3])}.${match[2]}`;
}

/** ISO YYYY-MM-DD → ДД.ММ.ГГГГ для поля «Срок выполнения». */
export function isoToDdMmYyyy(iso: string | null | undefined): string | null {
  const match = iso ? /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso) : null;
  if (!match) return null;
  return `${match[3]}.${match[2]}.${match[1]}`;
}

export function addDaysIso(iso: string, days: number): string {
  const date = new Date(`${iso}T12:00:00`);
  date.setDate(date.getDate() + days);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

/** GET/PUT /commercial/settings/promise-tracks-per-day — factory knob (T13). */
export type PromiseTracksPerDay = {
  tracks_per_day: number;
  updated_by: string | null;
  updated_at: string | null;
  min: number;
  max: number;
};

export const promiseKnobApi = {
  get: () => httpClient.get<PromiseTracksPerDay>(`${SETTINGS_BASE}/promise-tracks-per-day`),
  put: (tracksPerDay: number) =>
    httpClient.put<PromiseTracksPerDay>(
      `${SETTINGS_BASE}/promise-tracks-per-day`,
      JSON.stringify({ tracks_per_day: tracksPerDay }),
      JSON_HEADERS,
    ),
};

export const promiseKnobKeys = {
  all: ["promise-knob"] as const,
};

export const usePromiseKnobQuery = (enabled = true) =>
  useQuery<PromiseTracksPerDay>({
    queryKey: promiseKnobKeys.all,
    queryFn: () => promiseKnobApi.get(),
    enabled,
    staleTime: 30_000,
  });

export const useUpdatePromiseKnobMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (tracksPerDay: number) => promiseKnobApi.put(tracksPerDay),
    onSuccess: (row) => {
      queryClient.setQueryData(promiseKnobKeys.all, row);
      queryClient.invalidateQueries({ queryKey: promiseQuoteKeys.all });
    },
  });
};
