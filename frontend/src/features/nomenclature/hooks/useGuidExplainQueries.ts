import { useQuery, useQueryClient } from "@tanstack/react-query";
import { nomenclatureApi } from "@/features/nomenclature/api/nomenclatureApi";
import { nomenclatureKeys } from "@/features/nomenclature/hooks/useNomenclatureQueries";
import {
  GUID_SEARCH_MIN_CHARS,
  type GuidExplainResponse,
  type GuidExplainSource,
  type GuidSearchResponse,
} from "@/features/nomenclature/types/nomenclature";

/** Источник панели «Почему нет GUID» для экономиста («Завести в 1С»). */
export const economistGuidExplainSource: GuidExplainSource = {
  scope: "economist",
  explain: (productKind, mark) => nomenclatureApi.explainGuid(productKind, mark),
  search: (productKind, query) => nomenclatureApi.searchGuid(productKind, query),
  previewReplace: (body) => nomenclatureApi.previewGuidReplace(body),
  commitReplace: (body) => nomenclatureApi.commitGuidReplace(body),
};

export const useRefreshGuidTasks = () => {
  let queryClient: ReturnType<typeof useQueryClient> | null = null;
  try {
    queryClient = useQueryClient();
  } catch {
    queryClient = null;
  }
  return () => {
    if (!queryClient || !nomenclatureKeys) {
      return;
    }
    void queryClient.invalidateQueries({ queryKey: nomenclatureKeys.all });
    void queryClient.invalidateQueries({ queryKey: ["guid-explain"] });
  };
};

export const useGuidExplainQuery = (
  source: GuidExplainSource,
  productKind: string,
  mark: string,
) =>
  useQuery<GuidExplainResponse>({
    queryKey: ["guid-explain", source.scope, productKind, mark],
    queryFn: () => source.explain(productKind, mark),
    staleTime: 15_000,
  });

export const useGuidSearchQuery = (
  source: GuidExplainSource,
  productKind: string,
  query: string,
) => {
  const text = query.trim();
  return useQuery<GuidSearchResponse>({
    queryKey: ["guid-search", source.scope, productKind, text],
    queryFn: () => source.search(productKind, text),
    enabled: text.length >= GUID_SEARCH_MIN_CHARS,
    staleTime: 15_000,
  });
};
