import { useMemo } from "react";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import {
  buildPlateLineHighlightMap,
  DOBOR_MARKER_HIGHLIGHT_STYLE,
  PLATE_LINE_HIGHLIGHT_STYLES,
  splitLineByDoborMarker,
  type PlateLineHighlightKind,
} from "@/features/commercial-offer/lib/plateLineHighlights";
import { AutoResizeTextarea } from "@/shared/ui/Field";

type PlateListEditorProps = {
  draft: CommercialDraftDetails;
  value: string;
  onChange: (value: string) => void;
  minHeight?: number;
};

const LEGEND_ITEMS: Array<{ kind: PlateLineHighlightKind; label: string }> = [
  { kind: "correction", label: "Исправлено при распознавании" },
  { kind: "unparsed", label: "Не попало в расчёт" },
  { kind: "wide", label: "Шире стандартной" },
  { kind: "dobor", label: "Пара добора" },
];

export const PlateListEditor = ({ draft, value, onChange, minHeight = 440 }: PlateListEditorProps) => {
  const lines = value.split("\n");
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
            padding: "0.8rem 0.9rem",
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
                  minHeight: "1.5em",
                  background: style?.background ?? "transparent",
                  boxShadow: style ? `inset 3px 0 0 ${style.border}` : undefined,
                }}
              >
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
          }}
        />
      </div>
    </div>
  );
};
