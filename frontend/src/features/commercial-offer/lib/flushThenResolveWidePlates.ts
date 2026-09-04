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

/** Match live overlay ids to flushed server lines (optional ПБ / «п» / «Плиты»). */
export const normalizeLineKey = (line: string): string =>
  line
    .trim()
    .toLowerCase()
    .replace(/^плиты\s+/i, "")
    .replace(/^пб\s+/i, "")
    // Do not use \b — in JS it does not treat Cyrillic «п» as a word char.
    .replace(/-(\d+(?:[.,]\d+)?)п(?=\s|$)/gi, "-$1")
    .replace(/\s+/g, " ");

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

export const buildWidePlateResolveDecisions = ({
  liveLines,
  flushedWideLines,
  decisionsById,
}: {
  liveLines: Array<Pick<LiveWidePlateLine, "id" | "line">>;
  flushedWideLines: Array<Pick<WidePlateLine, "id" | "line">>;
  decisionsById: Record<string, WidePlateDecisionRecord>;
}): WidePlateResolveDecision[] => {
  const liveByKey = new Map(liveLines.map((item) => [normalizeLineKey(item.line), item]));
  return flushedWideLines.map((flushed) => {
    const live = liveByKey.get(normalizeLineKey(flushed.line));
    const decision = decisionsById[flushed.id] ?? (live ? decisionsById[live.id] : undefined);
    return {
      lineId: flushed.id,
      sourceLine: flushed.line,
      action: decision?.action ?? "confirm",
      replacementText: decision?.replacementText ?? "",
    };
  });
};

export const flushThenResolveWidePlates = async ({
  draftId,
  flushText,
  persistedText,
  liveLines,
  decisionsById,
  currentWideLines = [],
  updateInput,
  resolveWidePlates,
}: {
  draftId: string;
  flushText: string;
  persistedText: string;
  liveLines: LiveWidePlateLine[];
  decisionsById: Record<string, WidePlateDecisionRecord>;
  currentWideLines?: WidePlateLine[];
  updateInput: (payload: FlushInputPayload) => Promise<CommercialDraftDetails>;
  resolveWidePlates: (payload: ResolveWidePlatesPayload) => Promise<unknown>;
}): Promise<void> => {
  let flushedWideLines = currentWideLines;
  if (flushText.trim() !== persistedText.trim()) {
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
