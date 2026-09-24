import { httpClient, resolveApiUrl } from "@/shared/api/httpClient";
import type {
  ArchiveFileKind,
  ArchiveOfferDetails,
  ArchiveOfferListItem,
  ArchiveProductTypeFilter,
  ArchiveSearchApiResponse,
  ArchiveSection,
  KpReadinessPositionsResponse,
  ProductionEstimate,
} from "@/features/commercial-archive/types/archive";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const BASE = "/api/v1/commercial/archive";

export const archiveApi = {
  list: (section: ArchiveSection, productType?: ArchiveProductTypeFilter) => {
    const params = new URLSearchParams({ section });
    if (productType && productType !== "all") {
      params.set("product_type", productType);
    }
    return httpClient.get<ArchiveOfferListItem[]>(`${BASE}?${params.toString()}`);
  },

  getById: (kpId: number) => httpClient.get<ArchiveOfferDetails>(`${BASE}/${kpId}`),

  getReadinessPositions: (kpId: number) =>
    httpClient.get<KpReadinessPositionsResponse>(`${BASE}/${kpId}/readiness/positions`),

  search: ({ kpId, customer }: { kpId?: number; customer?: string }) => {
    const params = new URLSearchParams();
    if (kpId !== undefined) {
      params.set("kp_id", String(kpId));
    } else if (customer !== undefined) {
      params.set("customer", customer);
    }
    return httpClient.get<ArchiveSearchApiResponse>(`${BASE}/search?${params.toString()}`);
  },

  updateDiscount: (kpId: number, discount: number) =>
    httpClient.patch<ArchiveOfferDetails>(
      `${BASE}/${kpId}/discount`,
      JSON.stringify({ discount }),
      { "Content-Type": "application/json" },
    ),

  bindCounterparty: (kpId: number, counterpartyId: number) =>
    httpClient.patch<ArchiveOfferDetails>(
      `${BASE}/${kpId}/counterparty`,
      JSON.stringify({ counterparty_id: counterpartyId }),
      { "Content-Type": "application/json" },
    ),

  createAndBindCounterparty: (
    kpId: number,
    payload: { name: string; code_1c: string; inn?: string | null; kpp?: string | null },
  ) =>
    httpClient.post<ArchiveOfferDetails>(
      `${BASE}/${kpId}/counterparty`,
      JSON.stringify({
        name: payload.name,
        code_1c: payload.code_1c,
        inn: payload.inn,
        kpp: payload.kpp,
      }),
      { "Content-Type": "application/json" },
    ),

  updateLogisticsCost: (
    kpId: number,
    payload: {
      logisticsCost: number;
      pileLogisticsCost?: number;
      pileTripOverrides?: Record<string, number>;
      longPileDelivery?: Record<string, { trip_cost: number }>;
    },
  ) =>
    httpClient.patch<ArchiveOfferDetails>(
      `${BASE}/${kpId}/logistics-cost`,
      JSON.stringify({
        logistics_cost: payload.logisticsCost,
        pile_logistics_cost: payload.pileLogisticsCost,
        pile_trip_overrides: payload.pileTripOverrides,
        long_pile_delivery: payload.longPileDelivery,
      }),
      { "Content-Type": "application/json" },
    ),

  delete: (kpId: number) => httpClient.delete<void>(`${BASE}/${kpId}`),

  moveToProduction: (kpId: number, executionTerms: string) =>
    httpClient.post<ArchiveOfferDetails>(
      `${BASE}/${kpId}/move-to-production`,
      JSON.stringify({ execution_terms: executionTerms }),
      { "Content-Type": "application/json" },
    ),

  getProductionEstimate: (kpId: number) =>
    httpClient.get<ProductionEstimate>(`${BASE}/${kpId}/production-estimate`),

  buildDocumentUrl: (kpId: number, kind: ArchiveFileKind): string =>
    resolveApiUrl(`${BASE}/${kpId}/files/${kind}`),

  downloadDocument: (kpId: number, kind: ArchiveFileKind) => {
    const fallbackName =
      kind === "schema"
        ? `КП_${kpId}_schema.pdf`
        : kind === "pdf"
          ? `КП_${kpId}.pdf`
          : kind === "xlsx_delivery_in_unit"
            ? `КП_${kpId}_с_доставкой_в_цене.xlsx`
            : `КП_${kpId}.xlsx`;
    return httpClient.download(`${BASE}/${kpId}/files/${kind}`, fallbackName);
  },

  buildCurrentPlanUrl: (): string => resolveApiUrl(`${BASE}/current-plan/gantt`),

  resume: (kpId: number) =>
    httpClient.post<CommercialDraftDetails>(`${BASE}/${kpId}/resume`),
};
