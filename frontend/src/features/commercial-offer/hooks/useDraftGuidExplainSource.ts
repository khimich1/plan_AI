import { useMemo } from "react";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import type { GuidExplainSource } from "@/features/nomenclature/types/nomenclature";

/** Источник панели «Почему нет GUID» для менеджера: ответы только по его черновику. */
export const useDraftGuidExplainSource = (draftId: string): GuidExplainSource =>
  useMemo(
    () => ({
      scope: `draft-${draftId}`,
      explain: (productKind, mark) => commercialOfferApi.explainGuid(draftId, productKind, mark),
      search: (productKind, query) => commercialOfferApi.searchGuid(draftId, productKind, query),
      previewReplace: (body) => commercialOfferApi.previewGuidReplace(draftId, body),
      commitReplace: async (body) => {
        await commercialOfferApi.replaceGuid(draftId, body);
        return { written: [], skipped: [], queue_notice: null };
      },
    }),
    [draftId],
  );
