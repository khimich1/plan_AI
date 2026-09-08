import { getProductTypeConfig } from "@/features/commercial-offer/lib/productTypeConfig";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

/** Number of source batches on the draft for the active product type. */
export const getDraftBatchCount = (draft: CommercialDraftDetails): number => {
  const field = getProductTypeConfig(draft.metadata.product_type).batchesField;
  return draft.metadata[field]?.length ?? 0;
};
