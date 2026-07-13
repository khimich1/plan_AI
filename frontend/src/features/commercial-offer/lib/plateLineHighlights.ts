import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

export type PlateLineHighlightKind = "correction" | "unparsed" | "wide";

export type PlateLineHighlight = {
  kind: PlateLineHighlightKind;
  title: string;
};

const normalizeLineKey = (line: string) => line.trim().toLowerCase();

export const buildPlateLineHighlightMap = (
  draft: CommercialDraftDetails,
  textLines: string[],
): Map<number, PlateLineHighlight> => {
  const highlights = new Map<number, PlateLineHighlight>();
  const lines = textLines;

  const unparsedKeys = new Set(
    (draft.metadata.unparsed_lines ?? []).map((line) => normalizeLineKey(line)),
  );
  const wideKeys = new Set(
    (draft.metadata.wide_plate_lines ?? []).map((item) => normalizeLineKey(item.line)),
  );

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
    if (unparsedKeys.has(key)) {
      highlights.set(index, {
        kind: "unparsed",
        title: "Строка не попала в расчёт — проверьте вручную",
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

export const PLATE_LINE_HIGHLIGHT_STYLES: Record<PlateLineHighlightKind, { background: string; border: string }> = {
  correction: { background: "#fffaeb", border: "#fec84b" },
  unparsed: { background: "#fff4ed", border: "#f9a86c" },
  wide: { background: "#fef3f2", border: "#f97066" },
};
