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
  SpecificationChoicePayload,
  SpecificationView,
} from "@/features/commercial-archive/types/archive";
import { saveBlobAs } from "@/shared/lib/downloadFile";
import type {
  SupplyContract,
  SupplyContractCreatePayload,
  SupplyContractPatchPayload,
  SupplyContractRegistryRow,
} from "@/features/commercial-archive/types/supplyContract";

export const archiveKeys = {
  all: ["archive"] as const,
  list: (section: ArchiveSection, productType: ArchiveProductTypeFilter = "all") =>
    ["archive", "list", section, productType] as const,
  detail: (kpId: number) => ["archive", "offer", kpId] as const,
  specification: (kpId: number) => ["archive", "specification", kpId] as const,
  readinessPositions: (kpId: number) => ["archive", "readiness-positions", kpId] as const,
  search: (state: ArchiveSearchState) =>
    [
      "archive",
      "search",
      state?.kind ?? null,
      state?.kind === "number" ? state.value : state?.kind === "customer" ? state.value : null,
    ] as const,
  estimate: (kpId: number) => ["archive", "estimate", kpId] as const,
  supplyContract: (kpId: number) => ["archive", "supply-contract", kpId] as const,
  supplyRegistry: () => ["archive", "supply-contracts"] as const,
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

export const useSpecificationQuery = (kpId: number | null) =>
  useQuery<SpecificationView>({
    queryKey: archiveKeys.specification(kpId ?? -1),
    queryFn: () => archiveApi.getSpecification(kpId as number),
    enabled: kpId !== null,
  });

export const useSaveSpecificationMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, choice }: { kpId: number; choice: SpecificationChoicePayload }) =>
      archiveApi.saveSpecification(kpId, choice),
    onSuccess: (view, { kpId }) => {
      queryClient.setQueryData(archiveKeys.specification(kpId), view);
      queryClient.invalidateQueries({ queryKey: archiveKeys.detail(kpId) });
    },
  });
};

export const useDownloadSpecificationMutation = () =>
  useMutation({
    mutationFn: async (input: number | { kpId: number; format?: "xlsx" | "pdf" }) => {
      const kpId = typeof input === "number" ? input : input.kpId;
      const format = typeof input === "number" ? "xlsx" : (input.format ?? "xlsx");
      const result = await archiveApi.downloadSpecification(kpId, format);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
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

export const useBindCounterpartyMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, counterpartyId }: { kpId: number; counterpartyId: number }) =>
      archiveApi.bindCounterparty(kpId, counterpartyId),
    onSuccess: (offer) => {
      queryClient.setQueryData(archiveKeys.detail(offer.kp_id), offer);
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useCreateAndBindCounterpartyMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({
      kpId,
      name,
      code_1c,
      inn,
      kpp,
    }: {
      kpId: number;
      name: string;
      code_1c: string;
      inn?: string | null;
      kpp?: string | null;
    }) => archiveApi.createAndBindCounterparty(kpId, { name, code_1c, inn, kpp }),
    onSuccess: (offer) => {
      queryClient.setQueryData(archiveKeys.detail(offer.kp_id), offer);
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useUpdateLogisticsCostMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({
      kpId,
      logisticsCost,
      pileLogisticsCost,
      pileTripOverrides,
      longPileDelivery,
    }: {
      kpId: number;
      logisticsCost: number;
      pileLogisticsCost?: number;
      pileTripOverrides?: Record<string, number>;
      longPileDelivery?: Record<string, { trip_cost: number }>;
    }) =>
      archiveApi.updateLogisticsCost(kpId, {
        logisticsCost,
        pileLogisticsCost,
        pileTripOverrides,
        longPileDelivery,
      }),
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

export const useSupplyContractQuery = (kpId: number | null, enabled: boolean) =>
  useQuery<SupplyContract | null>({
    queryKey: archiveKeys.supplyContract(kpId ?? -1),
    queryFn: () => archiveApi.getSupplyContract(kpId as number),
    enabled: enabled && kpId !== null,
  });

export const useSupplyContractRegistryQuery = (enabled: boolean) =>
  useQuery<SupplyContractRegistryRow[]>({
    queryKey: archiveKeys.supplyRegistry(),
    queryFn: () => archiveApi.listSupplyContracts(),
    enabled,
  });

export const useCreateSupplyContractMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, payload }: { kpId: number; payload: SupplyContractCreatePayload }) =>
      archiveApi.createSupplyContract(kpId, payload),
    onSuccess: (contract, { kpId }) => {
      queryClient.setQueryData(archiveKeys.supplyContract(kpId), contract);
      queryClient.invalidateQueries({ queryKey: archiveKeys.supplyRegistry() });
    },
  });
};

export const useParseSupplyContractMutation = () =>
  useMutation({
    mutationFn: ({ kpId, file }: { kpId: number; file: File }) =>
      archiveApi.parseSupplyContract(kpId, file),
  });

export const usePatchSupplyContractMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ contractId, payload }: { contractId: number; payload: SupplyContractPatchPayload }) =>
      archiveApi.patchSupplyContract(contractId, payload),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: archiveKeys.supplyRegistry() });
      queryClient.invalidateQueries({ queryKey: ["archive", "supply-contract"] });
    },
  });
};

export const useReplaceSupplyContractMutation = (kpId: number) => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ contractId, payload }: { contractId: number; payload: SupplyContractCreatePayload }) =>
      archiveApi.replaceSupplyContract(contractId, payload),
    onSuccess: (contract) => {
      queryClient.setQueryData(archiveKeys.supplyContract(kpId), contract);
      queryClient.invalidateQueries({ queryKey: archiveKeys.supplyRegistry() });
    },
  });
};

export const useDownloadSupplyContractMutation = () =>
  useMutation({
    mutationFn: async ({ kpId, format }: { kpId: number; format: "docx" | "pdf" }) => {
      const result = await archiveApi.downloadSupplyContract(kpId, format);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
  });

export const useDownloadEdoAgreementMutation = () =>
  useMutation({
    mutationFn: async ({ kpId, format }: { kpId: number; format: "docx" | "pdf" }) => {
      const result = await archiveApi.downloadEdoAgreement(kpId, format);
      saveBlobAs(result.blob, result.filename);
      return result;
    },
  });

export const useExportInvoiceMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, warehouse }: { kpId: number; warehouse: string }) =>
      archiveApi.exportInvoice(kpId, warehouse),
    onSuccess: (offer) => {
      queryClient.setQueryData(archiveKeys.detail(offer.kp_id), offer);
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useExportInvoiceCorrectionMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, warehouse }: { kpId: number; warehouse?: string }) =>
      archiveApi.exportInvoiceCorrection(kpId, warehouse),
    onSuccess: (offer) => {
      queryClient.setQueryData(archiveKeys.detail(offer.kp_id), offer);
      queryClient.invalidateQueries({ queryKey: archiveKeys.all });
    },
  });
};

export const useSetArchivePaymentMutation = () => {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ kpId, paid }: { kpId: number; paid: boolean }) =>
      archiveApi.setPayment(kpId, paid),
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
