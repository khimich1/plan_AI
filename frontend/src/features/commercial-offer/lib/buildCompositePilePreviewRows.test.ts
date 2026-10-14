import { describe, expect, it } from "vitest";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import {
  buildCompositePileLinesFromOrderData,
  buildCompositePilePreviewRows,
} from "@/features/commercial-offer/lib/buildCompositePilePreviewRows";

const baseDraft = (): CommercialDraftDetails =>
  ({
    draft_id: "d1",
    order_data: [
      {
        product_kind: "composite_pile",
        product_type: "composite_piles",
        mark: "С60.30-ВС.1",
        name: "С60.30-ВС.1",
        concrete_grade: "B25",
        qty: 5,
        unit_price: 130,
        line_total: 650,
        line_id: "l1",
      },
      {
        product_kind: "composite_pile",
        product_type: "composite_piles",
        mark: "С80.30-НС.1",
        name: "С80.30-НС.1",
        concrete_grade: "B25",
        qty: 5,
        unit_price: 230,
        line_total: 1150,
        line_id: "l2",
      },
    ],
    metadata: {
      product_type: "composite_piles",
      default_concrete_grade: "B25",
      current_step: "composite_piles",
    },
  }) as CommercialDraftDetails;

describe("buildCompositePilePreviewRows", () => {
  it("builds canon section rows with grade qty price", () => {
    const rows = buildCompositePilePreviewRows(baseDraft());
    expect(rows).toHaveLength(2);
    expect(rows[0].mark).toBe("С60.30-ВС.1");
    expect(rows[0].concrete_grade).toBe("B25");
    expect(rows[0].qty).toBe(5);
    expect(rows[0].unit_price).toBe(130);
    expect(rows[0].product_kind).toBe("composite_pile");
  });

  it("serializes lines for grades update", () => {
    const rows = buildCompositePilePreviewRows(baseDraft());
    expect(buildCompositePileLinesFromOrderData(rows)).toContain("С60.30-ВС.1 B25 5");
  });
});
