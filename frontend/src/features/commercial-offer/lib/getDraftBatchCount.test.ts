import { describe, expect, it } from "vitest";

import { getDraftBatchCount } from "@/features/commercial-offer/lib/getDraftBatchCount";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const makeDraft = (
  productType: string,
  meta: Record<string, unknown>,
): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    metadata: {
      product_type: productType,
      plate_batches: [],
      ...meta,
    },
  }) as CommercialDraftDetails;

describe("getDraftBatchCount", () => {
  // S17
  it("returns fbs_batches length for fbs drafts", () => {
    const draft = makeDraft("fbs", {
      fbs_batches: [{ batch_index: 0 }, { batch_index: 1 }],
      plate_batches: [],
    });
    expect(getDraftBatchCount(draft)).toBe(2);
  });

  it("returns bridge_pile_batches length for bridge_piles drafts", () => {
    const draft = makeDraft("bridge_piles", {
      bridge_pile_batches: [{ batch_index: 0 }, { batch_index: 1 }, { batch_index: 2 }],
      plate_batches: [{ batch_index: 0 }],
    });
    expect(getDraftBatchCount(draft)).toBe(3);
  });

  it("returns plate_batches for plates", () => {
    const draft = makeDraft("plates", {
      plate_batches: [{ batch_index: 0 }],
    });
    expect(getDraftBatchCount(draft)).toBe(1);
  });
});
