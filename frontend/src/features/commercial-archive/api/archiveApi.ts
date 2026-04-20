import { httpClient, resolveApiUrl } from "@/shared/api/httpClient";
import type {
  ArchiveFileKind,
  ArchiveOfferDetails,
  ArchiveOfferListItem,
  ArchiveSearchResponse,
  ArchiveSection,
  ProductionEstimate,
} from "@/features/commercial-archive/types/archive";

const BASE = "/api/v1/commercial/archive";

export const archiveApi = {
  list: (section: ArchiveSection) =>
    httpClient.get<ArchiveOfferListItem[]>(`${BASE}?section=${encodeURIComponent(section)}`),

  getById: (kpId: number) => httpClient.get<ArchiveOfferDetails>(`${BASE}/${kpId}`),

  searchByNumber: (kpId: number) =>
    httpClient.get<ArchiveSearchResponse>(`${BASE}/search?query=${encodeURIComponent(String(kpId))}`),

  updateDiscount: (kpId: number, discount: number) =>
    httpClient.patch<ArchiveOfferDetails>(
      `${BASE}/${kpId}/discount`,
      JSON.stringify({ discount }),
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

  buildCurrentPlanUrl: (): string => resolveApiUrl(`${BASE}/current-plan/gantt`),
};
