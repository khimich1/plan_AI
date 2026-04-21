import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import type {
  ArchiveOfferDetails,
  ArchiveOfferListItem,
  ArchiveSearchResponse,
  ArchiveSection,
  ProductionEstimate,
} from "@/features/commercial-archive/types/archive";

export const archiveKeys = {
  all: ["archive"] as const,
  list: (section: ArchiveSection) => ["archive", "list", section] as const,
  detail: (kpId: number) => ["archive", "offer", kpId] as const,
  search: (kpId: number | null) => ["archive", "search", kpId] as const,
  estimate: (kpId: number) => ["archive", "estimate", kpId] as const,
};

export const useArchiveListQuery = (section: ArchiveSection) =>
  useQuery<ArchiveOfferListItem[]>({
    queryKey: archiveKeys.list(section),
    queryFn: () => archiveApi.list(section),
    staleTime: 15_000,
  });

export const useArchiveOfferQuery = (kpId: number | null) =>
  useQuery<ArchiveOfferDetails>({
    queryKey: archiveKeys.detail(kpId ?? -1),
    queryFn: () => archiveApi.getById(kpId as number),
    enabled: kpId !== null,
  });

export const useArchiveSearchQuery = (kpId: number | null) =>
  useQuery<ArchiveSearchResponse>({
    queryKey: archiveKeys.search(kpId),
    queryFn: () => archiveApi.searchByNumber(kpId as number),
    enabled: kpId !== null,
  });

export const useProductionEstimateQuery = (kpId: number | null) =>
  useQuery<ProductionEstimate>({
    queryKey: archiveKeys.estimate(kpId ?? -1),
    queryFn: () => archiveApi.getProductionEstimate(kpId as number),
    enabled: kpId !== null,
    staleTime: 60_000,
  });

export const useUpdateDiscountMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, discount }: { kpId: number; discount: number }) =>
      archiveApi.updateDiscount(kpId, discount),
    onSuccess: (offer) => {
      queryClient.setQueryData(archiveKeys.detail(offer.kp_id), offer);
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useDeleteOfferMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (kpId: number) => archiveApi.delete(kpId),
    onSuccess: (_, kpId) => {
      queryClient.removeQueries({ queryKey: archiveKeys.detail(kpId) });
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useMoveToProductionMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, executionTerms }: { kpId: number; executionTerms: string }) =>
      archiveApi.moveToProduction(kpId, executionTerms),
    onSuccess: (offer) => {
      queryClient.setQueryData(archiveKeys.detail(offer.kp_id), offer);
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};
