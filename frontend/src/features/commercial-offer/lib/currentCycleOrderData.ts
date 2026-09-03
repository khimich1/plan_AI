import type { CommercialDraftDetails, ProductType } from "@/features/commercial-offer/types/commercialOffer";
import { resolveDraftProductType } from "@/features/commercial-offer/lib/wizardStepOrder";

/** Sealed lines (prior append cycles) carry append_batch_id after start_append_cycle / calculate. */
export const isSealedOrderLine = (item: Record<string, unknown>): boolean =>
  String(item.append_batch_id ?? "").trim().length > 0;

/**
 * Lines belonging to the in-progress input cycle only — exclude already-added (sealed) batches.
 * Used for grade re-ingest / batch text so sealed same-type lines are not duplicated via append.
 */
export const getCurrentCycleOrderData = (
  draft: CommercialDraftDetails,
  productType: ProductType,
): CommercialDraftDetails["order_data"] => {
  const expected = resolveDraftProductType(productType);
  const metaType = resolveDraftProductType(draft.metadata.product_type);

  return (draft.order_data ?? []).filter((item) => {
    if (isSealedOrderLine(item)) {
      return false;
    }
    const rawType = String(item.product_type ?? "").trim().toLowerCase();
    if (!rawType) {
      // Legacy mono lines without product_type: treat as current cycle of metadata type.
      return metaType === expected;
    }
    return rawType === expected;
  });
};

/**
 * Full KP composition for one product type: sealed ∪ current-cycle lines of that type.
 * Used by step-1 preview panels so «Добавить к списку» never hides already-added lines.
 */
export const getProductTypeOrderData = (
  draft: CommercialDraftDetails,
  productType: ProductType,
): CommercialDraftDetails["order_data"] => {
  const expected = resolveDraftProductType(productType);
  const metaType = resolveDraftProductType(draft.metadata.product_type);

  return (draft.order_data ?? []).filter((item) => {
    const rawType = String(item.product_type ?? "").trim().toLowerCase();
    if (!rawType) {
      return metaType === expected;
    }
    return rawType === expected;
  });
};
