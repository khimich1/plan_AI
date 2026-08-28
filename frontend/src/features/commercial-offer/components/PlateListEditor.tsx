import { useMemo } from "react";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import {
  buildPlateLineHighlightMap,
  DOBOR_MARKER_HIGHLIGHT_STYLE,
  PLATE_LINE_HIGHLIGHT_STYLES,
  splitLineByDoborMarker,
  type PlateLineHighlightKind,
} from "@/features/commercial-offer/lib/plateLineHighlights";
import {
  assignNonEmptyLineNumbers,
  formatPlateLineNumber,
  lineNumberGutterCh,
} from "@/features/commercial-offer/lib/plateListLineNumbers";
import { AutoResizeTextarea } from "@/shared/ui/Field";

type PlateListEditorProps = {
  draft: CommercialDraftDetails;
  value: string;
  onChange: (value: string) => void;
  minHeight?: number;
  showLineNumbers?: boolean;
};

const EDITOR_PADDING_Y = "0.8rem";
const EDITOR_PADDING_X = "0.9rem";
const LINE_NUMBER_COLOR = "#667085";

const LEGEND_ITEMS: Array<{ kind: PlateLineHighlightKind; label: string }> = [
  { kind: "correction", label: "Исправлено при распознавании" },
  { kind: "unparsed", label: "Не попало в расчёт" },
  { kind: "wide", label: "Шире стандартной" },
  { kind: "dobor", label: "Пара добора" },
];

export const PlateListEditor = ({
  draft,
  value,
  onChange,
  minHeight = 440,
  showLineNumbers = false,
}: PlateListEditorProps) => {
  const lines = value.split("\n");
  const lineNumbers = showLineNumbers ? assignNonEmptyLineNumbers(lines) : [];
  const maxLineNumber = lineNumbers.reduce<number>(
    (max, number) => (number != null && number > max ? number : max),
    0,
  );
  const gutterCh = showLineNumbers ? lineNumberGutterCh(maxLineNumber) : 0;
  const textareaPaddingLeft = showLineNumbers
    ? `calc(${EDITOR_PADDING_X} + ${gutterCh}ch)`
    : EDITOR_PADDING_X;
  const highlightMap = useMemo(() => buildPlateLineHighlightMap(draft, lines), [draft, lines]);
  const activeKinds = useMemo(() => {
    const kinds = new Set<PlateLineHighlightKind>();
    lines.forEach((line, index) => {
      const trimmed = line.trim();
      if (!trimmed) {
        return;
      }
      const highlight = highlightMap.get(index);
      if (highlight) {
        kinds.add(highlight.kind);
      }
    });
    return kinds;
  }, [highlightMap, lines]);

  return (
    <div style={{ display: "grid", gap: "0.75rem" }}>
      {activeKinds.size > 0 && (
        <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap", fontSize: "0.85rem", color: "#475467" }}>
          {LEGEND_ITEMS.filter((item) => activeKinds.has(item.kind)).map((item) => (
            <span key={item.kind} style={{ display: "inline-flex", alignItems: "center", gap: "0.35rem" }}>
              <span
                style={{
                  width: 12,
                  height: 12,
                  borderRadius: 3,
                  background: PLATE_LINE_HIGHLIGHT_STYLES[item.kind].background,
                  border: `1px solid ${PLATE_LINE_HIGHLIGHT_STYLES[item.kind].border}`,
                }}
              />
              {item.label}
            </span>
          ))}
        </div>
      )}

      <div style={{ position: "relative", minHeight }}>
        <div
          aria-hidden
          style={{
            position: "absolute",
            inset: 0,
            overflow: "hidden",
            pointerEvents: "none",
            borderRadius: 12,
            border: "1px solid transparent",
            padding: `${EDITOR_PADDING_Y} ${EDITOR_PADDING_X}`,
            fontFamily: "Consolas, monospace",
            fontSize: "inherit",
            lineHeight: 1.5,
            whiteSpace: "pre-wrap",
            wordBreak: "break-word",
            color: "transparent",
          }}
        >
          {lines.map((line, index) => {
            const highlight = highlightMap.get(index);
            const style = highlight ? PLATE_LINE_HIGHLIGHT_STYLES[highlight.kind] : null;
            const lineNumber = showLineNumbers ? (lineNumbers[index] ?? null) : null;
            const doborBadge =
              highlight?.kind === "dobor" && highlight.doborPairId ? (
                <span
                  style={{
                    float: "right",
                    fontSize: "0.75rem",
                    color: "#026aa2",
                    fontWeight: 600,
                    letterSpacing: "0.02em",
                  }}
                >
                  ↔ добор
                </span>
              ) : null;
            return (
              <div
                key={`${index}-${line}`}
                title={highlight?.title}
                style={{
                  display: showLineNumbers ? "flex" : undefined,
                  alignItems: showLineNumbers ? "flex-start" : undefined,
                  minHeight: "1.5em",
                  background: style?.background ?? "transparent",
                  boxShadow: style ? `inset 3px 0 0 ${style.border}` : undefined,
                }}
              >
                {showLineNumbers && (
                  <span
                    data-testid={lineNumber != null ? "plate-line-number" : undefined}
                    aria-hidden
                    style={{
                      width: `${gutterCh}ch`,
                      flexShrink: 0,
                      color: LINE_NUMBER_COLOR,
                      userSelect: "none",
                      textAlign: "right",
                    }}
                  >
                    {lineNumber != null ? formatPlateLineNumber(lineNumber) : ""}
                  </span>
                )}
                <div style={showLineNumbers ? { flex: 1, minWidth: 0 } : undefined}>
                  {line ? (
                    splitLineByDoborMarker(line).map((segment, segmentIndex) =>
                      segment.isMarker ? (
                        <span
                          key={segmentIndex}
                          style={{
                            background: DOBOR_MARKER_HIGHLIGHT_STYLE.background,
                            boxShadow: `inset 0 -2px 0 ${DOBOR_MARKER_HIGHLIGHT_STYLE.border}`,
                          }}
                        >
                          {segment.text}
                        </span>
                      ) : (
                        <span key={segmentIndex}>{segment.text}</span>
                      ),
                    )
                  ) : (
                    "\u00a0"
                  )}
                  {doborBadge}
                </div>
              </div>
            );
          })}
        </div>

        <AutoResizeTextarea
          value={value}
          onChange={(event) => onChange(event.target.value)}
          placeholder="Пока нет списка плит."
          style={{
            minHeight,
            position: "relative",
            zIndex: 1,
            background: "transparent",
            lineHeight: 1.5,
            caretColor: "#101828",
            paddingLeft: textareaPaddingLeft,
          }}
        />
      </div>
    </div>
  );
};
