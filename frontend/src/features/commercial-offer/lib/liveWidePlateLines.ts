import type { CommercialDraftDetails, WidePlateLine } from "@/features/commercial-offer/types/commercialOffer";

export type LiveWidePlateLine = WidePlateLine;

const WIDE_MARK_RE =
  /^(?:ПБ\s*)?(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*-\s*([\d.,]+п?)(?:\s*шт\.?)?\s*(\d+)?(?:\s*шт\.?)?\s*$/i;

const collapseLineWhitespace = (line: string): string =>
  line.replace(/[\t\u00a0]+/g, " ").replace(/ {2,}/g, " ").trim();

const parseWidthDm = (raw: string): number => {
  const token = raw.replace(",", ".");
  const value = Number(token);
  if (!Number.isFinite(value)) {
    return 0;
  }
  if (token.includes(".") && value <= 2) {
    return value * 10;
  }
  return value;
};

/** Compare pasted vs parser-normalized wide marks (optional ПБ / п / qty). */
export const widePlateMatchKey = (line: string): string =>
  collapseLineWhitespace(line)
    .toLowerCase()
    .replace(/^плиты\s+/i, "")
    .replace(/^пб\s*/i, "")
    .replace(/-(\d+(?:[.,]\d+)?)п(?=\s|$)/gi, "-$1")
    .replace(/\s+\d+\s*(?:шт\.?)?\s*$/i, "")
    .replace(/\s*шт\.?\s*$/i, "")
    .replace(/\s+/g, " ");

export const liveWidePlateLines = (text: string): LiveWidePlateLine[] => {
  const result: LiveWidePlateLine[] = [];
  const lines = text.split(/\r?\n/);
  lines.forEach((raw, index) => {
    const original = raw.trim();
    const line = collapseLineWhitespace(original);
    if (!line) {
      return;
    }
    const match = line.match(WIDE_MARK_RE);
    if (!match) {
      return;
    }
    const widthDm = parseWidthDm(match[2] ?? "");
    const qty = match[4] ? Number(match[4]) : 1;
    if (!Number.isFinite(widthDm) || widthDm <= 12 || !Number.isFinite(qty) || qty <= 0) {
      return;
    }
    result.push({
      id: `live-wide-${index}`,
      line: original,
      qty,
    });
  });
  return result;
};

const textMatchKeys = (text: string): Set<string> => {
  const keys = new Set<string>();
  for (const raw of text.split(/\r?\n/)) {
    const key = widePlateMatchKey(raw);
    if (key) {
      keys.add(key);
    }
  }
  return keys;
};

export const overlayDraftWithLiveWideLines = (
  draft: CommercialDraftDetails,
  text: string,
): CommercialDraftDetails => {
  const live = liveWidePlateLines(text);
  const server = draft.metadata.wide_plate_lines ?? [];
  const serverByKey = new Map<string, WidePlateLine>();
  for (const item of server) {
    const key = widePlateMatchKey(item.line);
    if (key && !serverByKey.has(key)) {
      serverByKey.set(key, item);
    }
  }

  const usedKeys = new Set<string>();
  const merged: WidePlateLine[] = [];
  let liveHasNewWide = false;
  for (const item of live) {
    const key = widePlateMatchKey(item.line);
    const serverHit = key ? serverByKey.get(key) : undefined;
    if (key && !serverHit) {
      liveHasNewWide = true;
    }
    merged.push(serverHit ? { ...serverHit, line: item.line, qty: item.qty } : item);
    if (key) {
      usedKeys.add(key);
    }
  }

  const presentKeys = textMatchKeys(text);
  for (const item of server) {
    const key = widePlateMatchKey(item.line);
    if (!key || usedKeys.has(key) || !presentKeys.has(key)) {
      continue;
    }
    merged.push(item);
    usedKeys.add(key);
  }

  return {
    ...draft,
    metadata: {
      ...draft.metadata,
      wide_plate_lines: merged,
      wide_plates_resolved:
        merged.length === 0 || (Boolean(draft.metadata.wide_plates_resolved) && !liveHasNewWide),
    },
  };
};
