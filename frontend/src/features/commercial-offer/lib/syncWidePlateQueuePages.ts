import { normalizeLineKey } from "@/features/commercial-offer/lib/flushThenResolveWidePlates";
import type { WidePlateAction } from "@/features/commercial-offer/types/commercialOffer";

/**
 * Zip `metadata.plate_batches` onto recognition pages after wide-plate resolve.
 * Page i owns batch i (first photo is create, later photos append).
 * A length mismatch is a failure: callers must not invent page text.
 */

export type AppliedWideDecision = {
  sourceLine: string;
  action: WidePlateAction;
  replacementText: string;
};

/**
 * Unedited page takes its resolved batch.
 * An edited page keeps its own lines; the wide mark is replaced in full.
 */
export const resolveAppliedQueuePageText = ({
  pageText,
  batchBefore,
  batchAfter,
  decisions,
}: {
  pageText: string;
  batchBefore: string;
  batchAfter: string;
  decisions: readonly AppliedWideDecision[];
}): string => {
  if (pageText.trim() === batchBefore.trim()) {
    return batchAfter;
  }
  return applyWideDecisionsToPageText(pageText, decisions);
};

const replacementLines = (replacementText: string): string[] =>
  replacementText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

const applyWideDecisionsToPageText = (
  pageText: string,
  decisions: readonly AppliedWideDecision[],
): string => {
  const actionable = decisions.filter(
    (decision) => decision.action === "replace" || decision.action === "exclude",
  );
  const next: string[] = [];
  for (const line of pageText.split("\n")) {
    const key = normalizeLineKey(line);
    const decision = key
      ? actionable.find((item) => normalizeLineKey(item.sourceLine) === key)
      : undefined;
    if (!decision) {
      next.push(line);
      continue;
    }
    if (decision.action === "exclude") {
      continue;
    }
    const lines = replacementLines(decision.replacementText);
    next.push(...(lines.length > 0 ? lines : [line]));
  }
  return next.join("\n").trim();
};

/** Page i keeps its own resolved batch, or its edited text with the wide line replaced. */
export const applyWideResolveToQueuePages = ({
  pages,
  batchesBefore,
  resolvedPages,
  decisions,
}: {
  pages: readonly WidePlateQueuePage[];
  batchesBefore: readonly { normalized_text?: string | null }[];
  resolvedPages: readonly WidePlateQueuePage[];
  decisions: readonly AppliedWideDecision[];
}): WidePlateQueuePage[] =>
  resolvedPages.map((page, index) => {
    const previous = pages.find((item) => item.id === page.id);
    return {
      id: page.id,
      batchReviewText: resolveAppliedQueuePageText({
        pageText: previous?.batchReviewText ?? "",
        batchBefore: batchesBefore[index]?.normalized_text ?? "",
        batchAfter: page.batchReviewText,
        decisions,
      }),
    };
  });

export const WIDE_PLATE_QUEUE_LAYOUT_ERROR =
  "Число списков не совпадает с числом фото. Список у снимков не изменён.";

export type WidePlateQueuePage = {
  id: string;
  batchReviewText: string;
};

export type WidePlateQueueDraft = {
  metadata: {
    /** Joined draft text. Ignored: each page takes its own batch. */
    normalized_text?: string | null;
    plate_batches?: ReadonlyArray<{
      normalized_text?: string | null;
    }> | null;
  };
};

export type WidePlateQueueSyncResult =
  | { ok: true; pages: WidePlateQueuePage[]; activePageText: string }
  | { ok: false; message: string };

export const syncWidePlateQueuePages = ({
  pages,
  activePageId,
  draft,
}: {
  pages: readonly WidePlateQueuePage[];
  activePageId: string | null;
  draft: WidePlateQueueDraft;
}): WidePlateQueueSyncResult => {
  const batches = draft.metadata.plate_batches ?? [];
  if (batches.length !== pages.length) {
    return { ok: false, message: WIDE_PLATE_QUEUE_LAYOUT_ERROR };
  }
  const nextPages = pages.map((page, index) => ({
    id: page.id,
    batchReviewText: batches[index]?.normalized_text ?? "",
  }));
  const activePage = nextPages.find((page) => page.id === activePageId);
  if (!activePage) {
    return {
      ok: false,
      message: "Не удалось записать список к выбранному фото. Список у снимков не изменён.",
    };
  }
  return {
    ok: true,
    pages: nextPages,
    activePageText: activePage.batchReviewText,
  };
};
