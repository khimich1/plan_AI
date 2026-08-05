import type { CommercialDraftDetails, StepOrderLine } from "@/features/commercial-offer/types/commercialOffer";
import { toNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";

export const buildStepPreviewRows = (draft: CommercialDraftDetails): StepOrderLine[] =>
  (draft.order_data ?? []).map((item) => {
    const mark = String(item.mark ?? item.name ?? "").trim();
    const qty = toNumber(item.qty) ?? 0;
    const unitPrice = toNumber(item.unit_price);
    const lineTotal =
      unitPrice !== null && qty > 0 ? Number((unitPrice * qty).toFixed(2)) : toNumber(item.line_total);
    return {
      mark,
      name: String(item.name ?? mark),
      qty,
      unit_price: unitPrice,
      line_total: lineTotal,
      product_kind: "step",
    };
  });

export const buildStepLinesFromOrderData = (rows: StepOrderLine[]): string =>
  rows
    .filter((row) => row.mark && row.qty > 0)
    .map((row) => `${row.mark} ${row.qty}`)
    .join("\n");
