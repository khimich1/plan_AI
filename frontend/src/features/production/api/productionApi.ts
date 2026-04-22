import { httpClient } from "@/shared/api/httpClient";
import type {
  BuildPlanRequest,
  BuildPlanResponse,
  DayOccupancyResponse,
  DayViewResponse,
  DeletePlanResponse,
  GlobalCalendarResponse,
  KpCandidatesResponse,
  PlansMetadataResponse,
  WorkCalendarPayload,
} from "@/features/production/types/production";

const BASE = "/api/v1/production";

export const productionApi = {
  listPlans: () => httpClient.get<PlansMetadataResponse>(`${BASE}/plans`),

  getPlan: (planId: string) =>
    httpClient.get<Record<string, unknown>>(`${BASE}/plans/${encodeURIComponent(planId)}`),

  activatePlan: (planId: string) =>
    httpClient.post<{ plan_id: string; active: boolean }>(
      `${BASE}/plans/${encodeURIComponent(planId)}/activate`,
    ),

  deletePlan: (planId: string) =>
    httpClient.delete<DeletePlanResponse>(`${BASE}/plans/${encodeURIComponent(planId)}`),

  buildPlan: (payload: BuildPlanRequest) =>
    httpClient.post<BuildPlanResponse>(
      `${BASE}/plans/build`,
      JSON.stringify(payload),
      { "Content-Type": "application/json" },
    ),

  getCalendar: () => httpClient.get<GlobalCalendarResponse>(`${BASE}/calendar`),

  getDayOccupancy: (excludePlanId?: string | null) => {
    const query = excludePlanId
      ? `?exclude_plan_id=${encodeURIComponent(excludePlanId)}`
      : "";
    return httpClient.get<DayOccupancyResponse>(`${BASE}/day-occupancy${query}`);
  },

  listKpCandidates: () => httpClient.get<KpCandidatesResponse>(`${BASE}/kp-candidates`),

  getDayView: (date: string) =>
    httpClient.get<DayViewResponse>(`${BASE}/days/${encodeURIComponent(date)}`),

  completeDay: (date: string, planId: string) =>
    httpClient.post<{ plan_id: string; date: string; completed: boolean }>(
      `${BASE}/days/${encodeURIComponent(date)}/complete`,
      JSON.stringify({ plan_id: planId }),
      { "Content-Type": "application/json" },
    ),

  downloadDaySchema: (date: string) =>
    httpClient.download(
      `${BASE}/days/${encodeURIComponent(date)}/documents/schema`,
      `Схема_${date}.pdf`,
    ),

  downloadDayBreakdown: (date: string) =>
    httpClient.download(
      `${BASE}/days/${encodeURIComponent(date)}/documents/breakdown`,
      `Детальная_разбивка_${date}.xlsx`,
    ),

  downloadDayFormovka: (date: string) =>
    httpClient.download(
      `${BASE}/days/${encodeURIComponent(date)}/documents/formovka`,
      `Формовка_${date}.zip`,
    ),

  getWorkCalendar: () => httpClient.get<WorkCalendarPayload>(`${BASE}/work-calendar`),

  saveWorkCalendar: (payload: WorkCalendarPayload) =>
    httpClient.put<WorkCalendarPayload>(
      `${BASE}/work-calendar`,
      JSON.stringify(payload),
      { "Content-Type": "application/json" },
    ),
};
