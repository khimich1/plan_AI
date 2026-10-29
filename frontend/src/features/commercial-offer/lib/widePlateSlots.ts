import { liveWidePlateLines } from "@/features/commercial-offer/lib/liveWidePlateLines";
import { buildAutoSplitSuggestion } from "@/features/commercial-offer/lib/widePlateSuggestion";

export type WideSlotAction = "confirm" | "replace" | "exclude";

export type WideSlot = {
  id: string;
  lineIndex: number;
  action: WideSlotAction;
  replacementText: string;
  accepted: boolean;
};

const wideLineIndices = (lines: readonly string[]): Set<number> => {
  const indices = new Set<number>();
  for (const item of liveWidePlateLines(lines.join("\n"))) {
    const match = /^live-wide-(\d+)$/.exec(item.id);
    if (!match) {
      continue;
    }
    indices.add(Number(match[1]));
  }
  return indices;
};

const newWideSlot = (lineIndex: number): WideSlot => ({
  id: `wide-slot-${crypto.randomUUID()}`,
  lineIndex,
  action: "confirm",
  replacementText: "",
  accepted: false,
});

const exactLineMatches = (
  previousLines: readonly string[],
  nextLines: readonly string[],
): Array<{ previousIndex: number; nextIndex: number }> => {
  const previousCount = previousLines.length;
  const nextCount = nextLines.length;
  const scores = Array.from({ length: previousCount + 1 }, () => Array<number>(nextCount + 1).fill(0));

  for (let previousIndex = previousCount - 1; previousIndex >= 0; previousIndex -= 1) {
    for (let nextIndex = nextCount - 1; nextIndex >= 0; nextIndex -= 1) {
      scores[previousIndex]![nextIndex] =
        previousLines[previousIndex] === nextLines[nextIndex]
          ? scores[previousIndex + 1]![nextIndex + 1]! + 1
          : Math.max(scores[previousIndex + 1]![nextIndex]!, scores[previousIndex]![nextIndex + 1]!);
    }
  }

  const matches: Array<{ previousIndex: number; nextIndex: number }> = [];
  let previousIndex = 0;
  let nextIndex = 0;
  while (previousIndex < previousCount && nextIndex < nextCount) {
    if (previousLines[previousIndex] === nextLines[nextIndex]) {
      matches.push({ previousIndex, nextIndex });
      previousIndex += 1;
      nextIndex += 1;
      continue;
    }
    if (scores[previousIndex + 1]![nextIndex]! >= scores[previousIndex]![nextIndex + 1]!) {
      previousIndex += 1;
    } else {
      nextIndex += 1;
    }
  }
  return matches;
};

const indicesBetween = (start: number, end: number): number[] => {
  const indices: number[] = [];
  for (let index = start + 1; index < end; index += 1) {
    indices.push(index);
  }
  return indices;
};

/**
 * Length changed: matched lines keep their slot in order.
 * Exactly one deleted and one inserted line between matches is the same slot edited.
 * Any other gap drops the old slot; new wide lines get new slots.
 */
const carryWhenLineCountChanges = (
  previousLines: readonly string[],
  nextLines: readonly string[],
  slots: readonly WideSlot[],
): WideSlot[] => {
  const wide = wideLineIndices(nextLines);
  const slotsByIndex = new Map(slots.map((slot) => [slot.lineIndex, slot]));
  const anchors = [
    { previousIndex: -1, nextIndex: -1 },
    ...exactLineMatches(previousLines, nextLines),
    { previousIndex: previousLines.length, nextIndex: nextLines.length },
  ];
  const carried: WideSlot[] = [];

  const pushSlot = (lineIndex: number, slot: WideSlot | undefined, textChanged: boolean) => {
    if (!wide.has(lineIndex)) {
      return;
    }
    if (!slot) {
      carried.push(newWideSlot(lineIndex));
      return;
    }
    carried.push({
      ...slot,
      lineIndex,
      accepted: textChanged ? false : slot.accepted,
    });
  };

  for (let anchorIndex = 0; anchorIndex < anchors.length - 1; anchorIndex += 1) {
    const left = anchors[anchorIndex];
    const right = anchors[anchorIndex + 1];
    if (!left || !right) {
      continue;
    }
    const deleted = indicesBetween(left.previousIndex, right.previousIndex);
    const inserted = indicesBetween(left.nextIndex, right.nextIndex);

    if (deleted.length === 1 && inserted.length === 1) {
      const previousIndex = deleted[0]!;
      const nextIndex = inserted[0]!;
      pushSlot(
        nextIndex,
        slotsByIndex.get(previousIndex),
        previousLines[previousIndex] !== nextLines[nextIndex],
      );
    } else {
      for (const nextIndex of inserted) {
        pushSlot(nextIndex, undefined, true);
      }
    }

    if (right.previousIndex < previousLines.length) {
      pushSlot(right.nextIndex, slotsByIndex.get(right.previousIndex), false);
    }
  }

  return carried;
};

/** Same list length keeps each slot on its index. A line that is no longer wide loses the slot. A new wide line gets a slot. */
export const carryWideSlots = (
  previousLines: readonly string[],
  nextLines: readonly string[],
  slots: readonly WideSlot[],
): WideSlot[] => {
  if (previousLines.length === nextLines.length) {
    const wide = wideLineIndices(nextLines);
    const slotsByIndex = new Map(slots.map((slot) => [slot.lineIndex, slot]));
    const carried: WideSlot[] = [];
    const indices = [...wide].sort((left, right) => left - right);
    for (const lineIndex of indices) {
      const slot = slotsByIndex.get(lineIndex);
      if (!slot) {
        carried.push(newWideSlot(lineIndex));
        continue;
      }
      const textChanged = previousLines[lineIndex] !== nextLines[lineIndex];
      carried.push({ ...slot, accepted: textChanged ? false : slot.accepted });
    }
    return carried;
  }
  return carryWhenLineCountChanges(previousLines, nextLines, slots);
};

/** One card per line. An open slot stays visible; among open copies the filled choice wins. */
export const uniqueWideSlotsByLine = (slots: readonly WideSlot[]): WideSlot[] => {
  const rank = (slot: WideSlot): number =>
    (slot.accepted ? 0 : 8) + (slot.action !== "confirm" ? 2 : 0) + (slot.replacementText.trim() ? 1 : 0);
  const byLine = new Map<number, WideSlot>();
  for (const slot of slots) {
    const current = byLine.get(slot.lineIndex);
    if (!current || rank(slot) > rank(current)) {
      byLine.set(slot.lineIndex, slot);
    }
  }
  return [...byLine.values()].sort((left, right) => left.lineIndex - right.lineIndex);
};

const replacementLines = (replacementText: string): string[] =>
  replacementText
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

/** Confirm keeps the lines. Exclude deletes the indexed line. Replace writes the field. */
export const applyWideSlot = (
  lines: readonly string[],
  index: number,
  decision: Pick<WideSlot, "action" | "replacementText" | "accepted">,
): { ok: true; lines: string[] } | { ok: false; error: string } => {
  if (decision.action === "confirm") {
    return { ok: true, lines: [...lines] };
  }
  if (decision.action === "exclude") {
    return { ok: true, lines: lines.filter((_, lineIndex) => lineIndex !== index) };
  }
  const replacement = replacementLines(decision.replacementText);
  if (replacement.length === 0) {
    return { ok: false, error: "Укажите заменяющие позиции." };
  }
  return {
    ok: true,
    lines: [...lines.slice(0, index), ...replacement, ...lines.slice(index + 1)],
  };
};

export type WideSlotCommitResult = {
  lines: string[];
  slots: WideSlot[];
  failedId: string | null;
  error: string | null;
};

/** Writes selected open slots from top to bottom. Other slots stay in the carry so a line does not grow a second card. */
export const commitWideSlots = (
  lines: readonly string[],
  slots: readonly WideSlot[],
  onlyIds?: readonly string[],
): WideSlotCommitResult => {
  const allowed = onlyIds ? new Set(onlyIds) : null;
  const ordered = [...slots].sort((left, right) => left.lineIndex - right.lineIndex);
  let nextLines = [...lines];
  let current = ordered.map((slot) => ({ ...slot }));

  for (const original of ordered) {
    if (original.accepted || (allowed && !allowed.has(original.id))) {
      continue;
    }
    const slot = current.find((item) => item.id === original.id);
    if (!slot || slot.accepted) {
      continue;
    }
    const applied = applyWideSlot(nextLines, slot.lineIndex, slot);
    if (!applied.ok) {
      return { lines: nextLines, slots: current, failedId: slot.id, error: applied.error };
    }
    const marked = current.map((item) => (item.id === slot.id ? { ...item, accepted: true } : item));
    current = carryWideSlots(nextLines, applied.lines, marked);
    nextLines = applied.lines;
  }

  return { lines: nextLines, slots: current, failedId: null, error: null };
};

/** Fills the split suggestion once, the first time replace is chosen on an empty field. */
export const chooseWideSlotAction = (
  slot: WideSlot,
  action: WideSlotAction,
  line: string,
  qty: number,
): WideSlot => {
  if (action !== "replace" || slot.replacementText.trim() !== "") {
    return { ...slot, action };
  }
  return {
    ...slot,
    action,
    replacementText: buildAutoSplitSuggestion(line, qty),
  };
};

export type TextReviewConfirmPlan =
  | { closeReview: false; text: string; slots: WideSlot[]; error: string }
  | { closeReview: true; text: string; slots: WideSlot[]; uploadText: string | null };

/** Writes open text-review slots before «Список верен» compares the list with the server batch. */
export const prepareTextReviewConfirm = (
  editorText: string,
  slots: readonly WideSlot[],
  serverBatchText: string,
): TextReviewConfirmPlan => {
  const committed = commitWideSlots(editorText.split(/\r?\n/), slots);
  const text = committed.lines.join("\n");
  if (committed.error) {
    return { closeReview: false, text, slots: committed.slots, error: committed.error };
  }
  const trimmed = text.trim();
  const uploadText = trimmed && trimmed !== serverBatchText.trim() ? trimmed : null;
  return { closeReview: true, text, slots: committed.slots, uploadText };
};

/** Slots for text review without a photo. Not a photo page id. */
export const TEXT_REVIEW_SLOT_PAGE_ID = "text-review";

type WidePlateApplyRoute = {
  textReview?: boolean;
};

/** Open photo review and open text review write locally. A confirmed photo queue still resolves. */
export const routeWidePlateApply = (
  photoQueue: boolean,
  route?: WidePlateApplyRoute,
): "local" | "resolve" => {
  if (!photoQueue && route?.textReview) {
    return "local";
  }
  return photoQueue ? "local" : "resolve";
};

/** Photo queue: an open slot, or a wide line with no accepted slot, blocks the step. */
export const photoQueueWidePlateBlocksFinish = (
  pages: ReadonlyArray<{ text: string; slots: readonly WideSlot[] }>,
): boolean =>
  pages.some((page) => {
    const wide = wideLineIndices(page.text.split(/\r?\n/));
    for (const lineIndex of wide) {
      const slot = page.slots.find((item) => item.lineIndex === lineIndex);
      if (!slot?.accepted) {
        return true;
      }
    }
    return false;
  });
