import type { CommercialDraftDetails, PileOrderLine } from "@/features/commercial-offer/types/commercialOffer";
import {
  getProductTypeOrderData,
  isSealedOrderLine,
} from "@/features/commercial-offer/lib/currentCycleOrderData";
import { formatLineSourceText } from "@/features/commercial-offer/lib/formatLineSourceText";
import { toNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";

export type CompositePileOrderLine = Omit<PileOrderLine, "product_kind"> & {
  product_kind?: "composite_pile";
};

export const buildCompositePilePreviewRows = (
  draft: CommercialDraftDetails,
): CompositePileOrderLine[] =>
  getProductTypeOrderData(draft, "composite_piles").map((item) => {
    const mark = String(item.mark ?? item.name ?? "").trim();
    const qty = toNumber(item.qty) ?? 0;
    const unitPrice = toNumber(item.unit_price);
    const lineTotal =
      unitPrice !== null && qty > 0 ? Number((unitPrice * qty).toFixed(2)) : toNumber(item.line_total);
    return {
      lineId: typeof item.line_id === "string" && item.line_id.trim() ? item.line_id : null,
      sourceText: formatLineSourceText(item),
      mark,
      name: String(item.name ?? mark),
      concrete_grade: String(item.concrete_grade ?? draft.metadata.default_concrete_grade ?? "B25"),
      qty,
      unit_price: unitPrice,
      line_total: lineTotal,
      product_kind: "composite_pile",
      sealed: isSealedOrderLine(item),
    };
  });

export const buildCompositePileLinesFromOrderData = (rows: CompositePileOrderLine[]): string =>
  rows
    .filter((row) => row.mark && row.qty > 0 && !row.sealed)
    .map((row) => `${row.mark} ${row.concrete_grade} ${row.qty}`)
    .join("\n");
