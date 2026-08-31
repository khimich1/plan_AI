import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { archiveKeys } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { productionApi } from "@/features/production/api/productionApi";
import { saveBlobAs } from "@/shared/lib/downloadFile";
import { ApiError } from "@/shared/lib/apiError";
import { handlePlanVersionConflict } from "@/shared/lib/planConflict";
import type {
  AnalyzeSubstratesRequest,
  AnalyzeSubstratesResponse,
  BuildPlanRequest,
  BuildPlanResponse,
  DayCapacityMapResponse,
  DayDocumentKind,
  DayOccupancyResponse,
  DayViewResponse,
  GlobalCalendarResponse,
  KpCandidatesResponse,
  PlanDetailResponse,
  PlansMetadataResponse,
  RejectedPlateItem,
  RemoveTrackResponse,
  SaveDayCapacityRequest,
  SaveDayCapacityResponse,
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
  dayCapacity: (from: string, to: string) =>
    ["production", "day-capacity", from, to] as const,
  dayCapacityRoot: () => ["production", "day-capacity"] as const,
};

export const usePlansListQuery = () =>
  useQuery<PlansMetadataResponse>({
    queryKey: productionKeys.plans(),
    queryFn: productionApi.listPlans,
    staleTime: 15_000,
  });

export const usePlanQuery = (planId: string | null) =>
  useQuery<PlanDetailResponse>({
    queryKey: productionKeys.plan(planId ?? ""),
    queryFn: () => productionApi.getPlan(planId as string),
    enabled: planId !== null,
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
    onError: async (error, planId) => {
      await handlePlanVersionConflict(queryClient, error, {
        variables: { planId },
      });
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
    onError: async (error, planId) => {
      await handlePlanVersionConflict(queryClient, error, {
        variables: { planId },
      });
    },
  });
};

export const useBuildPlanMutation = () => {
  const queryClient = useQueryClient();
  return useMutation<BuildPlanResponse, ApiError, BuildPlanRequest>({
    mutationFn: (payload) => productionApi.buildPlan(payload),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
    },
    onError: async (error, variables) => {
      await handlePlanVersionConflict(queryClient, error, { variables });
    },
  });
};

/** Read-only analysis: urgent positions + substrates + capacity deficit. */
export const useAnalyzeSubstratesMutation = () =>
  useMutation<AnalyzeSubstratesResponse, ApiError, AnalyzeSubstratesRequest>({
    mutationFn: (payload) => productionApi.analyzeSubstrates(payload),
  });


export const useCompleteDayMutation = () => {
  const queryClient = useQueryClient();
  return useMutation<
    Awaited<ReturnType<typeof productionApi.completeDay>>,
    ApiError,
    {
      date: string;
      planId: string;
      rejectedPlates?: RejectedPlateItem[];
      expectedVersion?: number;
    }
  >({
    mutationFn: ({ date, planId, rejectedPlates = [], expectedVersion }) =>
      productionApi.completeDay(date, planId, rejectedPlates, expectedVersion),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
    onError: async (error, variables) => {
      await handlePlanVersionConflict(queryClient, error, { variables });
    },
  });
};

export const useDeleteTrackMutation = () => {
  const queryClient = useQueryClient();
  return useMutation<
    RemoveTrackResponse,
    ApiError,
    { planId: string; date: string; trackIndex: number; expectedVersion?: number }
  >({
    mutationFn: ({ planId, date, trackIndex, expectedVersion }) =>
      productionApi.deleteTrack(planId, date, trackIndex, expectedVersion),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: productionKeys.all });
    },
    onError: async (error, variables) => {
      await handlePlanVersionConflict(queryClient, error, { variables });
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

export const usePlanSgpExportMutation = () =>
  useMutation({
    mutationKey: ["production", "sgp-export"],
    mutationFn: async (planId: string) => {
      const result = await productionApi.downloadPlanSgpExport(planId);
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
    onError: async (error) => {
      await handlePlanVersionConflict(queryClient, error);
    },
  });
};

export const useDayCapacityQuery = (from: string, to: string, enabled = true) =>
  useQuery<DayCapacityMapResponse>({
    queryKey: productionKeys.dayCapacity(from, to),
    queryFn: () => productionApi.getDayCapacity(from, to),
    enabled: enabled && Boolean(from) && Boolean(to),
    staleTime: 30_000,
  });

export const useSaveDayCapacityMutation = () => {
  const queryClient = useQueryClient();
  return useMutation<SaveDayCapacityResponse, ApiError, SaveDayCapacityRequest>({
    mutationFn: (payload) => productionApi.saveDayCapacity(payload),
    onSuccess: (data) => {
      queryClient.setQueriesData<DayCapacityMapResponse>(
        { queryKey: productionKeys.dayCapacityRoot() },
        (prev) => {
          if (!prev) return prev;
          return {
            capacity: { ...prev.capacity, [data.date]: data.max_tracks },
          };
        },
      );
      queryClient.invalidateQueries({ queryKey: productionKeys.dayCapacityRoot() });
      queryClient.invalidateQueries({ queryKey: productionKeys.calendar() });
    },
  });
};
