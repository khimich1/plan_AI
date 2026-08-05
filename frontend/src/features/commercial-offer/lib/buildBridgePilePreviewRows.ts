import type { CommercialDraftDetails, PileOrderLine } from "@/features/commercial-offer/types/commercialOffer";
import { toNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";

export type BridgePileOrderLine = Omit<PileOrderLine, "product_kind"> & {
  available_grades?: string[];
  product_kind?: "bridge_pile";
};

export const buildBridgePilePreviewRows = (draft: CommercialDraftDetails): BridgePileOrderLine[] =>
  (draft.order_data ?? []).map((item) => {
    const mark = String(item.mark ?? item.name ?? "").trim();
    const qty = toNumber(item.qty) ?? 0;
    const unitPrice = toNumber(item.unit_price);
    const lineTotal =
      unitPrice !== null && qty > 0 ? Number((unitPrice * qty).toFixed(2)) : toNumber(item.line_total);
    const available = Array.isArray(item.available_grades)
      ? (item.available_grades as string[])
      : undefined;
    return {
      mark,
      name: String(item.name ?? mark),
      concrete_grade: String(item.concrete_grade ?? draft.metadata.default_concrete_grade ?? "B25"),
      available_grades: available,
      qty,
      unit_price: unitPrice,
      line_total: lineTotal,
      product_kind: "bridge_pile",
    };
  });

export const buildBridgePileLinesFromOrderData = (rows: BridgePileOrderLine[]): string =>
  rows
    .filter((row) => row.mark && row.qty > 0)
    .map((row) => `${row.mark} ${row.concrete_grade} ${row.qty}`)
    .join("\n");
