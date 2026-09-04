import { useMemo } from "react";

import { useSourceTextLint } from "@/features/commercial-offer/hooks/useSourceTextLint";
import {
  mergeReviewHighlights,
  type PlateLineHighlight,
} from "@/features/commercial-offer/lib/plateLineHighlights";
import type { CommercialDraftDetails, ProductType } from "@/features/commercial-offer/types/commercialOffer";

type UseBatchReviewHighlightsArgs = {
  text: string;
  productType: ProductType;
  draft: CommercialDraftDetails | null;
  enabled: boolean;
};

/** Live lint + existing draft highlights on batch-review (red reject, yellow soft). */
export const useBatchReviewHighlights = ({
  text,
  productType,
  draft,
  enabled,
}: UseBatchReviewHighlightsArgs): Map<number, PlateLineHighlight> => {
  const lint = useSourceTextLint({ text, productType, enabled });
  return useMemo(
    () => mergeReviewHighlights(draft, text, enabled ? lint.lines : [], { batchReview: true }),
    [draft, text, enabled, lint.lines],
  );
};
