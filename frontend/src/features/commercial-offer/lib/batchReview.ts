import type { CommercialDraftDetails, PlateBatch } from "@/features/commercial-offer/types/commercialOffer";

const normalizeLineKey = (line: string) => line.trim().toLowerCase();

const batchLineKeys = (batchText: string): Set<string> => {
  const keys = new Set<string>();
  for (const line of batchText.split("\n")) {
    const key = normalizeLineKey(line);
    if (key) {
      keys.add(key);
    }
  }
  return keys;
};

export const getCurrentPlateBatch = (draft: CommercialDraftDetails | null): PlateBatch | null => {
  const batches = draft?.metadata.plate_batches ?? [];
  return batches.length > 0 ? batches[batches.length - 1]! : null;
};

/** Text shown in OCR side-by-side review — last batch only, fallback for legacy drafts. */
export const getCurrentBatchReviewText = (draft: CommercialDraftDetails | null): string => {
  const batch = getCurrentPlateBatch(draft);
  if (batch?.normalized_text?.trim()) {
    return batch.normalized_text;
  }
  return draft?.metadata.normalized_text ?? "";
};

export const needsBatchReview = (draft: CommercialDraftDetails | null, confirmedBatchCount: number): boolean => {
  const batchCount = draft?.metadata.plate_batches?.length ?? 0;
  return batchCount > confirmedBatchCount;
};

/** Rebuild full input text when the user edits the last batch before confirming. */
export const mergeEditedBatchIntoFullText = (batches: PlateBatch[], editedLastBatchText: string): string => {
  const trimmedEdit = editedLastBatchText.trim();
  if (batches.length === 0) {
    return trimmedEdit;
  }
  const parts = batches
    .slice(0, -1)
    .map((batch) => batch.normalized_text.trim())
    .filter(Boolean);
  if (trimmedEdit) {
    parts.push(trimmedEdit);
  }
  return parts.join("\n");
};

/**
 * Draft slice for PlateListEditor highlights — OCR metadata limited to current batch lines.
 * ocr_corrections from the server already reflect the latest recognition pass.
 */
export const filterDraftForBatchReview = (
  draft: CommercialDraftDetails,
  batchText: string,
): CommercialDraftDetails => {
  const keys = batchLineKeys(batchText);
  if (keys.size === 0) {
    return draft;
  }

  const unparsed_lines = (draft.metadata.unparsed_lines ?? []).filter((line) => keys.has(normalizeLineKey(line)));
  const wide_plate_lines = (draft.metadata.wide_plate_lines ?? []).filter((item) => keys.has(normalizeLineKey(item.line)));

  return {
    ...draft,
    metadata: {
      ...draft.metadata,
      unparsed_lines,
      wide_plate_lines,
    },
  };
};
