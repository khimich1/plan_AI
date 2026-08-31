import { resolveDraftProductType } from "@/features/commercial-offer/lib/wizardStepOrder";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

/** Number of source batches on the draft for the active product type. */
export const getDraftBatchCount = (draft: CommercialDraftDetails): number => {
  const productType = resolveDraftProductType(draft.metadata.product_type);
  if (productType === "piles") {
    return draft.metadata.pile_batches?.length ?? 0;
  }
  if (productType === "steps") {
    return draft.metadata.step_batches?.length ?? 0;
  }
  if (productType === "marches") {
    return draft.metadata.march_batches?.length ?? 0;
  }
  if (productType === "bridge_piles") {
    return draft.metadata.bridge_pile_batches?.length ?? 0;
  }
  if (productType === "fbs") {
    return draft.metadata.fbs_batches?.length ?? 0;
  }
  return draft.metadata.plate_batches?.length ?? 0;
};
