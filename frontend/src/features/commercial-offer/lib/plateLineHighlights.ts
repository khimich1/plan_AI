import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

export type PlateLineHighlightKind = "correction" | "unparsed" | "wide" | "invalid_width" | "dobor";

export type PlateLineHighlight = {
  kind: PlateLineHighlightKind;
  title: string;
  doborPairId?: string;
  doborPartnerLine?: string;
};

const normalizeLineKey = (line: string) => line.trim().toLowerCase();

const stripUnparsedSuffix = (line: string) => line.replace(/\s*\(пропущено:.*\)\s*$/i, "").trim();

export const buildPlateLineHighlightMap = (
  draft: CommercialDraftDetails,
  textLines: string[],
): Map<number, PlateLineHighlight> => {
  const highlights = new Map<number, PlateLineHighlight>();
  const lines = textLines;

  const unparsedKeys = new Set(
    (draft.metadata.unparsed_lines ?? []).map((line) => normalizeLineKey(stripUnparsedSuffix(line))),
  );
  const wideKeys = new Set(
    (draft.metadata.wide_plate_lines ?? []).map((item) => normalizeLineKey(item.line)),
  );
  const invalidWidthKeys = new Set(
    (draft.metadata.invalid_width_lines ?? []).flatMap((item) =>
      [item.line, item.name].filter(Boolean).map((value) => normalizeLineKey(value)),
    ),
  );

  const doborByLineKey = new Map<string, { pairId: string; partnerLine: string }>();
  for (const pair of draft.metadata.dobor_pairs ?? []) {
    const primaryKey = normalizeLineKey(pair.primary_line);
    const complementKey = normalizeLineKey(pair.complement_line);
    doborByLineKey.set(primaryKey, { pairId: pair.id, partnerLine: pair.complement_line });
    doborByLineKey.set(complementKey, { pairId: pair.id, partnerLine: pair.primary_line });
  }

  const correctionRows = new Set<number>();
  for (const correction of draft.metadata.ocr_corrections ?? []) {
    if (correction.action === "verify_failed") {
      continue;
    }
    if (correction.row_index != null && correction.row_index > 0) {
      correctionRows.add(correction.row_index - 1);
    }
  }

  lines.forEach((line, index) => {
    const key = normalizeLineKey(line);
    if (wideKeys.has(key)) {
      highlights.set(index, {
        kind: "wide",
        title: "Позиция шире стандартной — требует решения ниже",
      });
      return;
    }
    if (invalidWidthKeys.has(key) && !draft.metadata.invalid_widths_resolved) {
      highlights.set(index, {
        kind: "invalid_width",
        title: "Нестандартная ширина — решение ниже",
      });
      return;
    }
    if (unparsedKeys.has(key)) {
      highlights.set(index, {
        kind: "unparsed",
        title: "Строка не попала в расчёт — проверьте вручную",
      });
      return;
    }
    const doborMatch = doborByLineKey.get(key);
    if (doborMatch) {
      highlights.set(index, {
        kind: "dobor",
        title: `Добор: пара с «${doborMatch.partnerLine}»`,
        doborPairId: doborMatch.pairId,
        doborPartnerLine: doborMatch.partnerLine,
      });
      return;
    }
    if (correctionRows.has(index)) {
      highlights.set(index, {
        kind: "correction",
        title: "Строка исправлена при распознавании — сверьте с фото",
      });
    }
  });

  return highlights;
};

const DEFAULT_UNPARSED_TITLE = "Строка не попала в расчёт — проверьте вручную";

export const lintLinesToUnparsedHighlights = (
  lines: Array<{ index: number; empty: boolean; ok: boolean; reason_text: string | null }>,
): Map<number, PlateLineHighlight> => {
  const map = new Map<number, PlateLineHighlight>();
  for (const line of lines) {
    if (!line.empty && !line.ok) {
      map.set(line.index, {
        kind: "unparsed",
        title: line.reason_text || DEFAULT_UNPARSED_TITLE,
      });
    }
  }
  return map;
};

/**
 * Batch-review highlights: draft map (yellow OCR corrections, wide, dobor, unparsed)
 * plus live source lint. Parser-accepted lines (including 8н) drop stale unparsed.
 * Does not add a н→п / load-suffix heuristic.
 */
export const mergeReviewHighlights = (
  draft: CommercialDraftDetails | null,
  text: string,
  lintLines: Array<{ index: number; empty: boolean; ok: boolean; reason_text: string | null }>,
): Map<number, PlateLineHighlight> => {
  const lines = text.split("\n");
  const base = draft ? buildPlateLineHighlightMap(draft, lines) : new Map<number, PlateLineHighlight>();
  if (lintLines.length === 0) {
    return base;
  }
  const next = new Map(base);
  for (const line of lintLines) {
    if (line.empty) {
      continue;
    }
    if (!line.ok) {
      next.set(line.index, {
        kind: "unparsed",
        title: line.reason_text || DEFAULT_UNPARSED_TITLE,
      });
      continue;
    }
    if (next.get(line.index)?.kind === "unparsed") {
      next.delete(line.index);
    }
  }
  return next;
};

export const PLATE_LINE_HIGHLIGHT_STYLES: Record<PlateLineHighlightKind, { background: string; border: string }> = {
  correction: { background: "#fffaeb", border: "#fec84b" },
  unparsed: { background: "#fff4ed", border: "#f9a86c" },
  wide: { background: "#fef3f2", border: "#f97066" },
  invalid_width: { background: "#fef3f2", border: "#f97066" },
  dobor: { background: "#f0f9ff", border: "#36bffa" },
};

export const DOBOR_MARKER_HIGHLIGHT_STYLE = { background: "#d1fadf", border: "#32d583" };

export type DoborMarkerSegment = {
  text: string;
  isMarker: boolean;
};

/** Tail marker pattern aligned with core/dobor_split.py _DOБOR_MARKER_RE (marker text only). */
const DOBOR_MARKER_TAIL_RE =
  /(\s*(?:\+\s*)?(?:доб(?:ор)?))(?:\s*[-—]?\s*)?(?:(?:\d+)(?:\s*[-—]?\s*)?(?:шт\.?)?)?\s*$/iu;

export const splitLineByDoborMarker = (line: string): DoborMarkerSegment[] => {
  const match = line.match(DOBOR_MARKER_TAIL_RE);
  if (!match || match.index === undefined) {
    return [{ text: line, isMarker: false }];
  }

  const markerText = match[1] ?? "";
  const markerStart = match.index;
  const markerEnd = markerStart + markerText.length;
  const segments: DoborMarkerSegment[] = [];

  if (markerStart > 0) {
    segments.push({ text: line.slice(0, markerStart), isMarker: false });
  }
  segments.push({ text: markerText, isMarker: true });
  if (markerEnd < line.length) {
    segments.push({ text: line.slice(markerEnd), isMarker: false });
  }

  return segments;
};
