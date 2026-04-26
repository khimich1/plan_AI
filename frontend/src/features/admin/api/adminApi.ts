import { httpClient } from "@/shared/api/httpClient";
import type {
  DbResetReport,
  DbStatsResponse,
  RecoverPlatesResponse,
} from "@/features/admin/types/admin";

const BASE = "/api/v1/admin/db";

export const adminApi = {
  getStats: () => httpClient.get<DbStatsResponse>(`${BASE}/stats`),

  resetFull: () => httpClient.post<DbResetReport>(`${BASE}/reset/full`),

  resetKpOnly: () => httpClient.post<DbResetReport>(`${BASE}/reset/kp-only`),

  resetPlansOnly: () => httpClient.post<DbResetReport>(`${BASE}/reset/plans-only`),

  resetCalendarOnly: () =>
    httpClient.post<DbResetReport>(`${BASE}/reset/calendar-only`),

  recoverStuckPlates: () =>
    httpClient.post<RecoverPlatesResponse>(`${BASE}/recover-plates`),
};
