import type { CommercialDraftDetails, PileOrderLine } from "@/features/commercial-offer/types/commercialOffer";
import { getCurrentCycleOrderData } from "@/features/commercial-offer/lib/currentCycleOrderData";
import { formatLineSourceText } from "@/features/commercial-offer/lib/formatLineSourceText";
import { toNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";

export type FbsOrderLine = Omit<PileOrderLine, "product_kind"> & {
  available_grades?: string[];
  product_kind?: "fbs";
};

export const buildFbsPreviewRows = (draft: CommercialDraftDetails): FbsOrderLine[] =>
  getCurrentCycleOrderData(draft, "fbs").map((item) => {
    const mark = String(item.mark ?? item.name ?? "").trim();
    const qty = toNumber(item.qty) ?? 0;
    const unitPrice = toNumber(item.unit_price);
    const lineTotal =
      unitPrice !== null && qty > 0 ? Number((unitPrice * qty).toFixed(2)) : toNumber(item.line_total);
    const available = Array.isArray(item.available_grades)
      ? (item.available_grades as string[])
      : undefined;
    return {
      lineId: typeof item.line_id === "string" && item.line_id.trim() ? item.line_id : null,
      sourceText: formatLineSourceText(item),
      mark,
      name: String(item.name ?? mark),
      concrete_grade: String(item.concrete_grade ?? draft.metadata.default_concrete_grade ?? "B25"),
      available_grades: available,
      qty,
      unit_price: unitPrice,
      line_total: lineTotal,
      product_kind: "fbs",
    };
  });

export const buildFbsLinesFromOrderData = (rows: FbsOrderLine[]): string =>
  rows
    .filter((row) => row.mark && row.qty > 0)
    .map((row) => `${row.mark} ${row.concrete_grade} ${row.qty}`)
    .join("\n");
