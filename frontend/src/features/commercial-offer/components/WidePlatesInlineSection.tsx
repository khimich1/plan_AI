import { useEffect, useState } from "react";
import type { CommercialDraftDetails, WidePlateAction } from "@/features/commercial-offer/types/commercialOffer";
import { buildAutoSplitSuggestion } from "@/features/commercial-offer/lib/widePlateSuggestion";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Textarea } from "@/shared/ui/Field";

type WidePlateDecision = {
  action: WidePlateAction;
  replacementText: string;
};

type WidePlatesInlineSectionProps = {
  draft: CommercialDraftDetails;
  decisions: Record<string, WidePlateDecision>;
  isPending: boolean;
  errorMessage?: string | null;
  onDecisionChange: (lineId: string, action: WidePlateAction, replacementText: string) => void;
  onApply: () => void;
};

export const WidePlatesInlineSection = ({
  draft,
  decisions,
  isPending,
  errorMessage,
  onDecisionChange,
  onApply,
}: WidePlatesInlineSectionProps) => {
  const wideLines = draft.metadata.wide_plate_lines ?? [];
  const [expanded, setExpanded] = useState(false);
  const needsAttention = wideLines.length > 0 && !draft.metadata.wide_plates_resolved;

  useEffect(() => {
    wideLines.forEach((item) => {
      const currentDecision = decisions[item.id];
      if (currentDecision?.action === "replace" && !currentDecision.replacementText.trim()) {
        onDecisionChange(item.id, "replace", buildAutoSplitSuggestion(item.line, item.qty));
      }
    });
  }, [decisions, onDecisionChange, wideLines]);

  if (!needsAttention) {
    return null;
  }

  return (
    <Card
      title="Нестандартная ширина"
      subtitle="Плита шире стандартной — подтвердите, замените на рекомендуемые позиции или исключите из списка."
      actions={
        <button
          type="button"
          onClick={() => setExpanded((open) => !open)}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "0.35rem",
            border: "1px solid #fecdca",
            background: "#fef3f2",
            color: "#b42318",
            borderRadius: 999,
            padding: "0.2rem 0.65rem",
            fontSize: "0.8rem",
            fontWeight: 600,
            cursor: "pointer",
          }}
        >
          {expanded ? "▾" : "▸"} {wideLines.length} позиций требуют внимания
        </button>
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      {!expanded ? (
        <p style={{ margin: 0, color: "#475467" }}>
          Разверните блок, чтобы принять решение по каждой позиции, затем нажмите «Применить решения».
        </p>
      ) : (
        <div style={{ display: "grid", gap: "1rem" }}>
          {wideLines.map((item) => {
            const currentDecision = decisions[item.id] ?? { action: "confirm", replacementText: "" };
            const suggestedReplacement = buildAutoSplitSuggestion(item.line, item.qty);

            return (
              <div
                key={item.id}
                style={{
                  display: "grid",
                  gap: "0.75rem",
                  border: "1px solid #e4e7ec",
                  borderRadius: 12,
                  padding: "1rem",
                  background: "#fafafa",
                }}
              >
                <div>
                  <strong>{item.line}</strong>
                  <div style={{ color: "#475467", marginTop: "0.25rem" }}>Количество: {item.qty}</div>
                  <div style={{ color: "#b54708", marginTop: "0.35rem", fontSize: "0.9rem" }}>
                    Рекомендуем заменить на стандартные позиции или подтвердить как есть.
                  </div>
                </div>

                <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
                  {([
                    ["confirm", "Подтвердить"],
                    ["replace", "Заменить"],
                    ["exclude", "Исключить"],
                  ] as Array<[WidePlateAction, string]>).map(([value, label]) => (
                    <label
                      key={value}
                      style={{
                        display: "flex",
                        gap: "0.5rem",
                        alignItems: "center",
                        border: "1px solid #e4e7ec",
                        borderRadius: 12,
                        padding: "0.7rem 0.9rem",
                        background: "#ffffff",
                      }}
                    >
                      <input
                        type="radio"
                        checked={currentDecision.action === value}
                        onChange={() =>
                          onDecisionChange(
                            item.id,
                            value,
                            value === "replace" && !currentDecision.replacementText.trim()
                              ? suggestedReplacement
                              : currentDecision.replacementText,
                          )
                        }
                      />
                      <span>{label}</span>
                    </label>
                  ))}
                </div>

                {currentDecision.action === "replace" && (
                  <FieldWrapper
                    label="Заменяющие позиции"
                    hint="Каждая строка — отдельная позиция в списке."
                  >
                    <Textarea
                      value={currentDecision.replacementText}
                      onChange={(event) => onDecisionChange(item.id, "replace", event.target.value)}
                    />
                  </FieldWrapper>
                )}
              </div>
            );
          })}

          <Button type="button" onClick={onApply} disabled={isPending}>
            {isPending ? "Применяем..." : "Применить решения"}
          </Button>
        </div>
      )}
    </Card>
  );
};
