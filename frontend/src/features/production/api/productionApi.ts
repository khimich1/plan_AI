import { httpClient } from "@/shared/api/httpClient";
import type {
  BuildPlanRequest,
  BuildPlanResponse,
  CompleteDayResponse,
  DayOccupancyResponse,
  DayViewResponse,
  DeletePlanResponse,
  GlobalCalendarResponse,
  RemoveTrackResponse,
  KpCandidatesResponse,
  PlansMetadataResponse,
  PlanDetailResponse,
  RejectedPlateItem,
  WorkCalendarPayload,
} from "@/features/production/types/production";

const BASE = "/api/v1/production";

export const productionApi = {
  listPlans: () => httpClient.get<PlansMetadataResponse>(`${BASE}/plans`),

  getPlan: (planId: string) =>
    httpClient.get<PlanDetailResponse>(`${BASE}/plans/${encodeURIComponent(planId)}`),

  activatePlan: (planId: string) =>
    httpClient.post<{ plan_id: string; active: boolean }>(
      `${BASE}/plans/${encodeURIComponent(planId)}/activate`,
    ),

  deletePlan: (planId: string) =>
    httpClient.delete<DeletePlanResponse>(`${BASE}/plans/${encodeURIComponent(planId)}`),

  deleteTrack: (
    planId: string,
    date: string,
    trackIndex: number,
    expectedVersion?: number,
  ) => {
    const versionQuery =
      typeof expectedVersion === "number"
        ? `?expected_version=${encodeURIComponent(String(expectedVersion))}`
        : "";
    return httpClient.delete<RemoveTrackResponse>(
      `${BASE}/plans/${encodeURIComponent(planId)}/days/${encodeURIComponent(date)}/tracks/${trackIndex}${versionQuery}`,
    );
  },

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

  completeDay: (
    date: string,
    planId: string,
    rejectedPlates: RejectedPlateItem[] = [],
    expectedVersion?: number,
  ) =>
    httpClient.post<CompleteDayResponse>(
      `${BASE}/days/${encodeURIComponent(date)}/complete`,
      JSON.stringify({
        plan_id: planId,
        rejected_plates: rejectedPlates,
        ...(typeof expectedVersion === "number"
          ? { expected_version: expectedVersion }
          : {}),
      }),
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

  downloadPlanSgpExport: (planId: string) =>
    httpClient.download(
      `${BASE}/plans/${encodeURIComponent(planId)}/sgp-export`,
      `SGP_${planId}.xlsx`,
    ),

  getWorkCalendar: () => httpClient.get<WorkCalendarPayload>(`${BASE}/work-calendar`),

  saveWorkCalendar: (payload: WorkCalendarPayload) =>
    httpClient.put<WorkCalendarPayload>(
      `${BASE}/work-calendar`,
      JSON.stringify(payload),
      { "Content-Type": "application/json" },
    ),
};
