import type { CommercialDraftDetails, PileOrderLine } from "@/features/commercial-offer/types/commercialOffer";
import { getCurrentCycleOrderData } from "@/features/commercial-offer/lib/currentCycleOrderData";
import { toNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";

export const buildPilePreviewRows = (draft: CommercialDraftDetails): PileOrderLine[] =>
  getCurrentCycleOrderData(draft, "piles").map((item) => {
    const mark = String(item.mark ?? item.name ?? "").trim();
    const qty = toNumber(item.qty) ?? 0;
    const unitPrice = toNumber(item.unit_price);
    const lineTotal =
      unitPrice !== null && qty > 0 ? Number((unitPrice * qty).toFixed(2)) : toNumber(item.line_total);
    return {
      mark,
      name: String(item.name ?? mark),
      concrete_grade: String(item.concrete_grade ?? draft.metadata.default_concrete_grade ?? "B25"),
      qty,
      unit_price: unitPrice,
      line_total: lineTotal,
      product_kind: "pile",
    };
  });

export const buildPileLinesFromOrderData = (rows: PileOrderLine[]): string =>
  rows
    .filter((row) => row.mark && row.qty > 0)
    .map((row) => `${row.mark} ${row.concrete_grade} ${row.qty}`)
    .join("\n");
