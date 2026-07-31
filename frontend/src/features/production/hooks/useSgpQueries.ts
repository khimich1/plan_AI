import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import {
  sgpApi,
  type SgpFilter,
  type SgpPlatesResponse,
} from "@/features/production/api/sgpApi";
import { productionKeys } from "@/features/production/hooks/useProductionQueries";
import { archiveKeys } from "@/features/commercial-archive/hooks/useArchiveQueries";

export const sgpKeys = {
  all: ["sgp"] as const,
  plates: (filter: SgpFilter) => ["sgp", "plates", filter] as const,
  free: () => ["sgp", "free"] as const,
};

export const useSgpPlatesQuery = (filter: SgpFilter) =>
  useQuery<SgpPlatesResponse>({
    queryKey: sgpKeys.plates(filter),
    queryFn: () => sgpApi.listPlates(filter),
    staleTime: 10_000,
  });

export const useSgpFreePlatesQuery = (enabled = true) =>
  useQuery({
    queryKey: sgpKeys.free(),
    queryFn: () => sgpApi.freePlates(),
    enabled,
    staleTime: 10_000,
  });

export const useSgpUnlinkMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ sgpId, qty }: { sgpId: number; qty: number }) =>
      sgpApi.unlink(sgpId, qty),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: sgpKeys.all });
      void qc.invalidateQueries({ queryKey: productionKeys.all });
      void qc.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useSgpRelinkMutation = () => {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({
      sgpId,
      targetKpId,
      qty,
    }: {
      sgpId: number;
      targetKpId: number;
      qty: number;
    }) => sgpApi.relink(sgpId, targetKpId, qty),
    onSuccess: () => {
      void qc.invalidateQueries({ queryKey: sgpKeys.all });
      void qc.invalidateQueries({ queryKey: productionKeys.all });
      void qc.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};
