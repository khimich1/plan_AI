import { getCurrentBatch, getCurrentBatchReviewText } from "@/features/commercial-offer/lib/batchReview";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

export type ApplyAiSessionPlan = {
  /** When true, do not call resetSource() / multiPage.reset(). */
  preserveMultiSession: boolean;
  /** Text to write into the active page + store batch review field. */
  nextActivePageText: string | null;
};

/**
 * Text the review editor should show after Apply AI.
 * Prefers the AI batch, then draft normalized_text when the last OCR batch is stale.
 */
export const resolveApplyAiReviewText = (draft: CommercialDraftDetails): string => {
  const lastBatch = getCurrentBatch(draft);
  const lastBatchText = getCurrentBatchReviewText(draft);
  if (lastBatch?.source_type === "ai" && lastBatch.normalized_text?.trim()) {
    return lastBatch.normalized_text;
  }
  const normalized = draft.metadata.normalized_text ?? "";
  if (normalized.trim() && normalized.trim() !== lastBatchText.trim()) {
    return draft.metadata.normalized_text;
  }
  return lastBatchText;
};

/**
 * Post–Apply AI session plan (R10 + A.4).
 * Multi-page recognition must keep pages/statuses/hasStarted after AI.
 * Always returns the review text so the list beside the photo can sync.
 */
export const planApplyAiSessionSync = (args: {
  multiHasStarted: boolean;
  activePageId: string | null;
  draft: CommercialDraftDetails;
}): ApplyAiSessionPlan => {
  const reviewText = resolveApplyAiReviewText(args.draft);
  const nextActivePageText = reviewText.trim() ? reviewText : null;
  if (!args.multiHasStarted) {
    return { preserveMultiSession: false, nextActivePageText };
  }
  return {
    preserveMultiSession: true,
    nextActivePageText: args.activePageId ? nextActivePageText : null,
  };
};
