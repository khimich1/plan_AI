import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import {
  getProductTypeOrderData,
  isSealedOrderLine,
} from "@/features/commercial-offer/lib/currentCycleOrderData";
import { formatLineSourceText } from "@/features/commercial-offer/lib/formatLineSourceText";
import { toNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";

export type KpPreviewFlag = "wide_direct" | "wide_split";

export type KpPreviewRow = {
  lineId: string | null;
  name: string;
  qty: number;
  unitPrice: number | null;
  flag: KpPreviewFlag | null;
  sourceLine?: string;
  sourceText: string;
  sealed: boolean;
};

const WIDE_WIDTH_M = 1.2;
const DIMENSION_EPS = 0.01;
const WIDTH_SUM_EPS = 0.05;

type ParsedWideLine = {
  lengthM: number;
  widthM: number;
  rawLine: string;
};

export const parseWidePlateLine = (line: string): ParsedWideLine | null => {
  const normalized = line.trim().replace(",", ".");
  const lineMatch = normalized.match(/^(.*\S)\s+(\d+)$/);
  const platePart = (lineMatch?.[1] ?? normalized).trim();
  const nameMatch = platePart.match(/^(?:Плиты\s+)?(?:ПБ\s+)?([\d.]+)-([\d.]+)-/i);
  if (!nameMatch) {
    return null;
  }

  const lengthDm = Number(nameMatch[1]);
  const widthToken = nameMatch[2];
  const widthDm = Number(widthToken);
  if (!Number.isFinite(lengthDm) || !Number.isFinite(widthDm)) {
    return null;
  }

  const widthM = widthToken.includes(".") && widthDm <= 2 ? widthDm : widthDm / 10;

  return {
    lengthM: lengthDm / 10,
    widthM,
    rawLine: line,
  };
};

const findWideSourceLine = (
  wideLines: CommercialDraftDetails["metadata"]["wide_plate_lines"],
  lengthM: number,
): string | undefined => {
  for (const wide of wideLines) {
    const parsed = parseWidePlateLine(wide.line);
    if (parsed && Math.abs(parsed.lengthM - lengthM) < DIMENSION_EPS) {
      return wide.line;
    }
  }
  return undefined;
};

const buildWideSplitIndex = (
  orderData: CommercialDraftDetails["order_data"],
  wideLines: CommercialDraftDetails["metadata"]["wide_plate_lines"],
): Map<number, string> => {
  const splitSourceByIndex = new Map<number, string>();

  for (const wide of wideLines) {
    const parsed = parseWidePlateLine(wide.line);
    if (!parsed || parsed.widthM <= WIDE_WIDTH_M + DIMENSION_EPS) {
      continue;
    }

    const candidates: number[] = [];
    orderData.forEach((item, index) => {
      // Sealed (prior append) plates must not be flagged as wide_split for the current cycle.
      if (isSealedOrderLine(item)) {
        return;
      }
      const itemLengthM = toNumber(item.length_m);
      const itemWidthM = toNumber(item.width_m);
      if (itemLengthM === null || itemWidthM === null) {
        return;
      }
      if (Math.abs(itemLengthM - parsed.lengthM) < DIMENSION_EPS && itemWidthM <= WIDE_WIDTH_M + DIMENSION_EPS) {
        candidates.push(index);
      }
    });

    if (candidates.length === 0) {
      continue;
    }

    const sumWidths = candidates.reduce((acc, index) => acc + (toNumber(orderData[index].width_m) ?? 0), 0);
    const isSplit =
      candidates.length >= 2 || Math.abs(sumWidths - parsed.widthM) < WIDTH_SUM_EPS;

    if (!isSplit) {
      continue;
    }

    for (const index of candidates) {
      splitSourceByIndex.set(index, wide.line);
    }
  }

  return splitSourceByIndex;
};

export const buildKpPreviewRows = (draft: CommercialDraftDetails): KpPreviewRow[] => {
  const orderData = getProductTypeOrderData(draft, "plates");
  const wideLines = draft.metadata.wide_plate_lines ?? [];
  const widePlatesResolved = draft.metadata.wide_plates_resolved ?? false;
  const splitSourceByIndex =
    widePlatesResolved || wideLines.length === 0 ? new Map<number, string>() : buildWideSplitIndex(orderData, wideLines);

  return orderData.map((item, index) => {
    const sealed = isSealedOrderLine(item);
    const widthM = toNumber(item.width_m);
    let flag: KpPreviewRow["flag"] = null;
    let sourceLine: string | undefined;

    // Wide flags apply only to the in-progress cycle; sealed rows are display-only.
    if (!sealed && !widePlatesResolved) {
      if (widthM !== null && widthM > WIDE_WIDTH_M + DIMENSION_EPS) {
        flag = "wide_direct";
        const lengthM = toNumber(item.length_m);
        if (lengthM !== null) {
          sourceLine = findWideSourceLine(wideLines, lengthM);
        }
      } else if (splitSourceByIndex.has(index)) {
        flag = "wide_split";
        sourceLine = splitSourceByIndex.get(index);
      }
    }

    return {
      lineId: typeof item.line_id === "string" && item.line_id.trim() ? item.line_id : null,
      name: String(item.name ?? ""),
      qty: toNumber(item.qty) ?? 0,
      unitPrice: toNumber(item.unit_price),
      flag,
      sourceLine,
      sourceText: formatLineSourceText(item),
      sealed,
    };
  });
};
