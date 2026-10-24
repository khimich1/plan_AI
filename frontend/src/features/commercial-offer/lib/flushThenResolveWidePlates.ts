import type { LiveWidePlateLine } from "@/features/commercial-offer/lib/liveWidePlateLines";
import type {
  CommercialDraftDetails,
  WidePlateAction,
  WidePlateLine,
} from "@/features/commercial-offer/types/commercialOffer";

export type FlushPageSnapshot = {
  id: string;
  batchReviewText: string;
};

export type WidePlateDecisionRecord = {
  action: WidePlateAction;
  replacementText: string;
  /** Plate line the manager decided on. Lookup uses normalizeLineKey, not the row id. */
  sourceLine?: string;
};

export type WidePlateResolveDecision = {
  lineId?: string;
  sourceLine: string;
  action: WidePlateAction;
  replacementText: string;
};

export type FlushInputPayload = {
  draftId: string;
  text: string;
  image: null;
  mode: "replace";
};

export type ResolveWidePlatesPayload = {
  draftId: string;
  decisions: WidePlateResolveDecision[];
};

/** Same miss the server returns when a wide decision does not bind to a line. */
export const WIDE_PLATE_UNMATCHED_DECISION_ERROR =
  "Не удалось применить решение по широким плитам: позиция не найдена в списке.";

/**
 * Match live overlay lines to flushed server lines.
 * Mirrors `_normalize_wide_line_key`: case, «Плиты» / «ПБ», load suffix «п»,
 * «шт», and a trailing quantity. Does not change the displayed line.
 * Do not use `\b` — in JS it does not treat Cyrillic «п» as a word character.
 */
export const normalizeLineKey = (line: string): string => {
  let text = line.trim().toLowerCase();
  text = text.replace(/[\t\u00a0]+/g, " ");
  text = text.replace(/^плиты\s+/, "");
  text = text.replace(/^пб\s+/, "");
  text = text.replace(/-(\d+(?:[.,]\d+)?)п(?=\s|$)/g, "-$1");
  text = text.replace(/\s+/g, " ");
  text = text.replace(/\s+шт\.?\s*/g, " ");
  text = text.replace(/\s+\d+\s*$/, "");
  return text.trim();
};

export const buildMergedFlushText = ({
  hasStarted,
  pages,
  activePageId,
  editorText,
  singlePageText,
}: {
  hasStarted: boolean;
  pages: FlushPageSnapshot[];
  activePageId: string | null;
  editorText: string;
  singlePageText: string;
}): string => {
  if (!hasStarted || pages.length === 0) {
    return singlePageText.trim();
  }
  return pages
    .map((page) => (page.id === activePageId ? editorText.trim() : page.batchReviewText.trim()))
    .filter(Boolean)
    .join("\n");
};

/**
 * Text sent before wide-plate resolve.
 * A photo queue uses the list on each page. `storeText` is often the last
 * batch after hydrate-draft, while the screen still shows another photo.
 */
export const buildWidePlateApplyFlushText = ({
  hasStarted,
  pages,
  activePageId,
  storeText,
}: {
  hasStarted: boolean;
  pages: FlushPageSnapshot[];
  activePageId: string | null;
  storeText: string;
}): string => {
  if (!hasStarted || pages.length === 0) {
    return storeText.trim();
  }
  const activePage = pages.find((page) => page.id === activePageId);
  return buildMergedFlushText({
    hasStarted: true,
    pages,
    activePageId,
    editorText: activePage?.batchReviewText ?? storeText,
    singlePageText: storeText,
  });
};

const decisionMatchesMark = (decision: WidePlateDecisionRecord | undefined, key: string): boolean => {
  if (!decision) {
    return false;
  }
  if (!decision.sourceLine || !key) {
    return true;
  }
  return normalizeLineKey(decision.sourceLine) === key;
};

/** Id first, then the mark key, then any replace/exclude stored against the same plate mark. */
export const findWidePlateDecision = (
  item: { id: string; line: string },
  decisions: Record<string, WidePlateDecisionRecord>,
): WidePlateDecisionRecord | undefined => {
  const key = normalizeLineKey(item.line);
  const direct = decisions[item.id];
  if (decisionMatchesMark(direct, key)) {
    return direct;
  }
  const byMark = key ? decisions[key] : undefined;
  if (decisionMatchesMark(byMark, key)) {
    return byMark;
  }
  if (!key) {
    return undefined;
  }
  for (const decision of Object.values(decisions)) {
    if (decision.action !== "replace" && decision.action !== "exclude") {
      continue;
    }
    if (decision.sourceLine && normalizeLineKey(decision.sourceLine) === key) {
      return decision;
    }
  }
  return undefined;
};

export const widePlateDecisionOrConfirm = (
  item: { id: string; line: string },
  decisions: Record<string, WidePlateDecisionRecord>,
): WidePlateDecisionRecord =>
  findWidePlateDecision(item, decisions) ?? { action: "confirm", replacementText: "" };

const isReplaceOrExclude = (action: WidePlateAction | undefined): action is "replace" | "exclude" =>
  action === "replace" || action === "exclude";

/** Server already resolved, but the page still shows the mark and the manager chose replace or exclude. */
const decisionsFromPageWideLines = (
  liveLines: Array<Pick<LiveWidePlateLine, "id" | "line">>,
  decisionsById: Record<string, WidePlateDecisionRecord>,
): WidePlateResolveDecision[] => {
  const decisions: WidePlateResolveDecision[] = [];
  for (const live of liveLines) {
    const decision = findWidePlateDecision(live, decisionsById);
    if (!decision || !isReplaceOrExclude(decision.action)) {
      continue;
    }
    decisions.push({
      lineId: live.id,
      sourceLine: live.line,
      action: decision.action,
      replacementText: decision.replacementText,
    });
  }
  return decisions;
};

export const buildWidePlateResolveDecisions = ({
  liveLines,
  flushedWideLines,
  decisionsById,
}: {
  liveLines: Array<Pick<LiveWidePlateLine, "id" | "line">>;
  flushedWideLines: Array<Pick<WidePlateLine, "id" | "line">>;
  decisionsById: Record<string, WidePlateDecisionRecord>;
}): WidePlateResolveDecision[] => {
  if (flushedWideLines.length === 0) {
    return decisionsFromPageWideLines(liveLines, decisionsById);
  }
  const flushedKeys = new Set(flushedWideLines.map((item) => normalizeLineKey(item.line)));
  const unmatchedBinding = liveLines.some((live) => {
    const action = findWidePlateDecision(live, decisionsById)?.action;
    if (!isReplaceOrExclude(action)) {
      return false;
    }
    return !flushedKeys.has(normalizeLineKey(live.line));
  });
  if (unmatchedBinding) {
    throw new Error(WIDE_PLATE_UNMATCHED_DECISION_ERROR);
  }
  const liveByKey = new Map(liveLines.map((item) => [normalizeLineKey(item.line), item]));
  return flushedWideLines.map((flushed) => {
    const live = liveByKey.get(normalizeLineKey(flushed.line));
    const decision =
      (live ? findWidePlateDecision(live, decisionsById) : undefined) ??
      findWidePlateDecision({ id: flushed.id, line: flushed.line }, decisionsById);
    return {
      lineId: flushed.id,
      sourceLine: flushed.line,
      action: decision?.action ?? "confirm",
      replacementText: decision?.replacementText ?? "",
    };
  });
};

/** Drop replace/exclude whose mark is gone. An empty page list does not count as gone. */
export const pruneWidePlateActions = (
  actions: Record<string, WidePlateDecisionRecord>,
  pageText: string,
): Record<string, WidePlateDecisionRecord> => {
  if (!pageText.trim()) {
    return actions;
  }
  const present = new Set(
    pageText
      .split(/\r?\n/)
      .map((line) => normalizeLineKey(line))
      .filter(Boolean),
  );
  let changed = false;
  const next: Record<string, WidePlateDecisionRecord> = {};
  for (const [id, decision] of Object.entries(actions)) {
    const markLeft =
      isReplaceOrExclude(decision.action) &&
      Boolean(decision.sourceLine) &&
      !present.has(normalizeLineKey(decision.sourceLine ?? ""));
    if (markLeft) {
      changed = true;
      continue;
    }
    next[id] = decision;
  }
  return changed ? next : actions;
};

export const flushThenResolveWidePlates = async ({
  draftId,
  flushText,
  persistedText,
  liveLines,
  decisionsById,
  currentWideLines = [],
  photoQueue = false,
  updateInput,
  resolveWidePlates,
}: {
  draftId: string;
  flushText: string;
  persistedText: string;
  liveLines: LiveWidePlateLine[];
  decisionsById: Record<string, WidePlateDecisionRecord>;
  currentWideLines?: WidePlateLine[];
  /** Photo review must not collapse plate_batches via update mode replace. */
  photoQueue?: boolean;
  updateInput: (payload: FlushInputPayload) => Promise<CommercialDraftDetails>;
  resolveWidePlates: (payload: ResolveWidePlatesPayload) => Promise<unknown>;
}): Promise<void> => {
  let flushedWideLines = currentWideLines;
  if (!photoQueue && flushText.trim() !== persistedText.trim()) {
    const draft = await updateInput({
      draftId,
      text: flushText,
      image: null,
      mode: "replace",
    });
    flushedWideLines = draft.metadata.wide_plate_lines ?? [];
  }
  const decisions = buildWidePlateResolveDecisions({
    liveLines,
    flushedWideLines,
    decisionsById,
  });
  if (decisions.length === 0) {
    return;
  }
  await resolveWidePlates({ draftId, decisions });
};
