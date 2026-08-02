import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import type {
  ArchiveFileKind,
  ArchiveOfferDetails,
  ArchiveOfferListItem,
  ArchiveProductTypeFilter,
  ArchiveSearchApiResponse,
  ArchiveSearchState,
  ArchiveSection,
  KpReadinessPositionsResponse,
  ProductionEstimate,
} from "@/features/commercial-archive/types/archive";
import { saveBlobAs } from "@/shared/lib/downloadFile";

export const archiveKeys = {
  all: ["archive"] as const,
  list: (section: ArchiveSection, productType: ArchiveProductTypeFilter = "all") =>
    ["archive", "list", section, productType] as const,
  detail: (kpId: number) => ["archive", "offer", kpId] as const,
  readinessPositions: (kpId: number) => ["archive", "readiness-positions", kpId] as const,
  search: (state: ArchiveSearchState) =>
    [
      "archive",
      "search",
      state?.kind ?? null,
      state?.kind === "number" ? state.value : state?.kind === "customer" ? state.value : null,
    ] as const,
  estimate: (kpId: number) => ["archive", "estimate", kpId] as const,
};

export const useArchiveListQuery = (
  section: ArchiveSection,
  productTypeFilter: ArchiveProductTypeFilter = "all",
) =>
  useQuery<ArchiveOfferListItem[]>({
    queryKey: archiveKeys.list(section, productTypeFilter),
    queryFn: () => archiveApi.list(section, productTypeFilter),
    staleTime: 15_000,
  });

export const useArchiveOfferQuery = (kpId: number | null) =>
  useQuery<ArchiveOfferDetails>({
    queryKey: archiveKeys.detail(kpId ?? -1),
    queryFn: () => archiveApi.getById(kpId as number),
    enabled: kpId !== null,
  });

export const useKpReadinessPositionsQuery = (
  kpId: number | null,
  options?: { enabled?: boolean },
) =>
  useQuery<KpReadinessPositionsResponse>({
    queryKey: archiveKeys.readinessPositions(kpId ?? -1),
    queryFn: () => archiveApi.getReadinessPositions(kpId as number),
    enabled: kpId !== null && (options?.enabled ?? true),
    staleTime: 15_000,
  });

export const useArchiveSearchQuery = (searchState: ArchiveSearchState) =>
  useQuery<ArchiveSearchApiResponse>({
    queryKey: archiveKeys.search(searchState),
    queryFn: () => {
      if (searchState?.kind === "number") {
        return archiveApi.search({ kpId: searchState.value });
      }
      if (searchState?.kind === "customer") {
        return archiveApi.search({ customer: searchState.value });
      }
      throw new Error("Search state is empty");
    },
    enabled: searchState !== null,
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

export const useUpdateLogisticsCostMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, logisticsCost }: { kpId: number; logisticsCost: number }) =>
      archiveApi.updateLogisticsCost(kpId, logisticsCost),
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

export const useArchiveDocumentMutation = (kind: ArchiveFileKind) =>
  useMutation({
    mutationKey: ["archive", "document", kind],
    mutationFn: async (kpId: number) => {
      const result = await archiveApi.downloadDocument(kpId, kind);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
  });
