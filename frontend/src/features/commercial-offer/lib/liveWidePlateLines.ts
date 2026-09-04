import type { CommercialDraftDetails, WidePlateLine } from "@/features/commercial-offer/types/commercialOffer";

export type LiveWidePlateLine = WidePlateLine;

const WIDE_MARK_RE =
  /^(?:ПБ\s+)?(\d+(?:[.,]\d+)?)-(\d+(?:[.,]\d+)?)-(\S+)\s+(\d+)\s*$/i;

const parseWidthDm = (raw: string): number => Number(raw.replace(",", "."));

export const liveWidePlateLines = (text: string): LiveWidePlateLine[] => {
  const result: LiveWidePlateLine[] = [];
  const lines = text.split(/\r?\n/);
  lines.forEach((raw, index) => {
    const line = raw.trim();
    if (!line) {
      return;
    }
    const match = line.match(WIDE_MARK_RE);
    if (!match) {
      return;
    }
    const widthDm = parseWidthDm(match[2] ?? "");
    const qty = Number(match[4]);
    if (!Number.isFinite(widthDm) || widthDm <= 12 || !Number.isFinite(qty) || qty <= 0) {
      return;
    }
    result.push({
      id: `live-wide-${index}`,
      line,
      qty,
    });
  });
  return result;
};

export const overlayDraftWithLiveWideLines = (
  draft: CommercialDraftDetails,
  text: string,
): CommercialDraftDetails => {
  const wide_plate_lines = liveWidePlateLines(text);
  return {
    ...draft,
    metadata: {
      ...draft.metadata,
      wide_plate_lines,
      wide_plates_resolved: wide_plate_lines.length === 0 || Boolean(draft.metadata.wide_plates_resolved),
    },
  };
};
