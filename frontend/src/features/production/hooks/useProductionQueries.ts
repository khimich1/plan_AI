import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { archiveKeys } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { productionApi } from "@/features/production/api/productionApi";
import { saveBlobAs } from "@/shared/lib/downloadFile";
import type {
  BuildPlanRequest,
  BuildPlanResponse,
  DayDocumentKind,
  DayOccupancyResponse,
  DayViewResponse,
  GlobalCalendarResponse,
  KpCandidatesResponse,
  PlansMetadataResponse,
  RejectedPlateItem,
  RemoveTrackResponse,
  WorkCalendarPayload,
} from "@/features/production/types/production";

export const productionKeys = {
  all: ["production"] as const,
  plans: () => ["production", "plans"] as const,
  plan: (planId: string) => ["production", "plan", planId] as const,
  calendar: () => ["production", "calendar"] as const,
  day: (date: string) => ["production", "day", date] as const,
  occupancy: (excludePlanId?: string | null) =>
    ["production", "occupancy", excludePlanId ?? "all"] as const,
  kpCandidates: () => ["production", "kp-candidates"] as const,
  workCalendar: () => ["production", "work-calendar"] as const,
};

export const usePlansListQuery = () =>
  useQuery<PlansMetadataResponse>({
    queryKey: productionKeys.plans(),
    queryFn: productionApi.listPlans,
    staleTime: 15_000,
  });

export const useGlobalCalendarQuery = () =>
  useQuery<GlobalCalendarResponse>({
    queryKey: productionKeys.calendar(),
    queryFn: productionApi.getCalendar,
    staleTime: 15_000,
  });

export const useDayViewQuery = (date: string | null) =>
  useQuery<DayViewResponse>({
    queryKey: productionKeys.day(date ?? ""),
    queryFn: () => productionApi.getDayView(date as string),
    enabled: date !== null,
    // День открывается редко и данные по нему критичны (плиты, документы):
    // всегда забираем свежие, чтобы Drawer не показывал stale-результат,
    // оставшийся после пересборки плана.
    staleTime: 0,
    gcTime: 0,
    refetchOnMount: "always",
  });

export const useDayOccupancyQuery = (excludePlanId?: string | null) =>
  useQuery<DayOccupancyResponse>({
    queryKey: productionKeys.occupancy(excludePlanId),
    queryFn: () => productionApi.getDayOccupancy(excludePlanId),
    staleTime: 15_000,
  });

export const useKpCandidatesQuery = (enabled = true) =>
  useQuery<KpCandidatesResponse>({
    queryKey: productionKeys.kpCandidates(),
    queryFn: productionApi.listKpCandidates,
    enabled,
    staleTime: 30_000,
  });

export const useWorkCalendarQuery = () =>
  useQuery<WorkCalendarPayload>({
    queryKey: productionKeys.workCalendar(),
    queryFn: productionApi.getWorkCalendar,
    staleTime: 60_000,
  });

export const useActivatePlanMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (planId: string) => productionApi.activatePlan(planId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
    },
  });
};

export const useDeletePlanMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (planId: string) => productionApi.deletePlan(planId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
    },
  });
};

export const useBuildPlanMutation = () => {
  const queryClient = useQueryClient();
  return useMutation<BuildPlanResponse, Error, BuildPlanRequest>({
    mutationFn: (payload) => productionApi.buildPlan(payload),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
    },
  });
};

export const useCompleteDayMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({
      date,
      planId,
      rejectedPlates = [],
    }: {
      date: string;
      planId: string;
      rejectedPlates?: RejectedPlateItem[];
    }) => productionApi.completeDay(date, planId, rejectedPlates),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useDeleteTrackMutation = () => {
  const queryClient = useQueryClient();
  return useMutation<
    RemoveTrackResponse,
    Error,
    { planId: string; date: string; trackIndex: number }
  >({
    mutationFn: ({ planId, date, trackIndex }) =>
      productionApi.deleteTrack(planId, date, trackIndex),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
    },
  });
};

export const useDayDocumentMutation = (kind: DayDocumentKind) =>
  useMutation({
    mutationKey: ["production", "day-document", kind],
    mutationFn: async (date: string) => {
      const downloader =
        kind === "schema"
          ? productionApi.downloadDaySchema
          : kind === "breakdown"
            ? productionApi.downloadDayBreakdown
            : productionApi.downloadDayFormovka;
      const result = await downloader(date);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
  });

export const useSaveWorkCalendarMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (payload: WorkCalendarPayload) => productionApi.saveWorkCalendar(payload),
    onSuccess: (data) => {
      queryClient.setQueryData(productionKeys.workCalendar(), data);
      queryClient.invalidateQueries({ queryKey: productionKeys.workCalendar() });
    },
  });
};
