import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { adminApi } from "@/features/admin/api/adminApi";
import { archiveKeys } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { productionKeys } from "@/features/production/hooks/useProductionQueries";
import type {
  DbResetReport,
  DbStatsResponse,
  RecoverPlatesResponse,
} from "@/features/admin/types/admin";

export const adminKeys = {
  all: ["admin"] as const,
  stats: ["admin", "stats"] as const,
};

const useInvalidateAfterReset = () => {
  const queryClient = useQueryClient();
  return () => {
    queryClient.invalidateQueries({ queryKey: adminKeys.all });
    queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    queryClient.invalidateQueries({ queryKey: productionKeys.all });
  };
};

export const useDbStatsQuery = (enabled: boolean = true) =>
  useQuery<DbStatsResponse>({
    queryKey: adminKeys.stats,
    queryFn: () => adminApi.getStats(),
    enabled,
    staleTime: 0,
    refetchOnWindowFocus: false,
  });

export const useFullResetMutation = () => {
  const invalidateAll = useInvalidateAfterReset();
  return useMutation<DbResetReport, Error, void>({
    mutationFn: () => adminApi.resetFull(),
    onSuccess: invalidateAll,
  });
};

export const useKpResetMutation = () => {
  const invalidateAll = useInvalidateAfterReset();
  return useMutation<DbResetReport, Error, void>({
    mutationFn: () => adminApi.resetKpOnly(),
    onSuccess: invalidateAll,
  });
};

export const usePlansResetMutation = () => {
  const invalidateAll = useInvalidateAfterReset();
  return useMutation<DbResetReport, Error, void>({
    mutationFn: () => adminApi.resetPlansOnly(),
    onSuccess: invalidateAll,
  });
};

export const useCalendarResetMutation = () => {
  const invalidateAll = useInvalidateAfterReset();
  return useMutation<DbResetReport, Error, void>({
    mutationFn: () => adminApi.resetCalendarOnly(),
    onSuccess: invalidateAll,
  });
};

export const useRecoverPlatesMutation = () => {
  const invalidateAll = useInvalidateAfterReset();
  return useMutation<RecoverPlatesResponse, Error, void>({
    mutationFn: () => adminApi.recoverStuckPlates(),
    onSuccess: invalidateAll,
  });
};
