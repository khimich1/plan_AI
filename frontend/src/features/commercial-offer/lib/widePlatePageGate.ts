import { liveWidePlateLines, widePlateMatchKey } from "@/features/commercial-offer/lib/liveWidePlateLines";

/** Wide-mark keys still written on a photo list. */
export const wideKeysInText = (text: string): string[] => {
  const keys: string[] = [];
  const seen = new Set<string>();
  for (const line of liveWidePlateLines(text)) {
    const key = widePlateMatchKey(line.line);
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    keys.push(key);
  }
  return keys;
};

/** Wide marks on this photo that this photo has not applied yet. */
export const undecidedWideKeys = (text: string, settledKeys: readonly string[]): string[] => {
  const settled = new Set(settledKeys);
  return wideKeysInText(text).filter((key) => !settled.has(key));
};

export const photoQueueHasUndecidedWidePlate = (
  pages: readonly { id: string; batchReviewText: string }[],
  settledByPage: Readonly<Record<string, readonly string[]>>,
): boolean =>
  pages.some(
    (page) => undecidedWideKeys(page.batchReviewText, settledByPage[page.id] ?? []).length > 0,
  );
