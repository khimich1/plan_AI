import { useEffect, useState } from "react";
import type {
  CommercialDraftDetails,
  InvalidWidthAction,
  InvalidWidthLine,
} from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

export type InvalidWidthDecisionState = {
  action: InvalidWidthAction;
  widthMm: number | null;
};

type InvalidWidthsInlineSectionProps = {
  draft: CommercialDraftDetails;
  decisions: Record<string, InvalidWidthDecisionState>;
  isPending: boolean;
  errorMessage?: string | null;
  onDecisionChange: (lineId: string, action: InvalidWidthAction, widthMm: number | null) => void;
  onApply: () => void;
};

const formatPrice = (price: number) =>
  new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 }).format(price);

export const defaultInvalidWidthDecision = (item: InvalidWidthLine): InvalidWidthDecisionState => {
  if (item.replacements.length === 0) {
    return { action: "exclude", widthMm: null };
  }
  const upper = item.replacements.reduce((best, repl) =>
    repl.width_mm > best.width_mm ? repl : best,
  );
  return { action: "replace_width", widthMm: upper.width_mm };
};

const choiceValue = (decision: InvalidWidthDecisionState): string => {
  if (decision.action === "exclude") {
    return "exclude";
  }
  return `width:${decision.widthMm ?? ""}`;
};

export const InvalidWidthsInlineSection = ({
  draft,
  decisions,
  isPending,
  errorMessage,
  onDecisionChange,
  onApply,
}: InvalidWidthsInlineSectionProps) => {
  const invalidLines = draft.metadata.invalid_width_lines ?? [];
  const [expanded, setExpanded] = useState(false);
  const needsAttention = invalidLines.length > 0 && !draft.metadata.invalid_widths_resolved;

  useEffect(() => {
    invalidLines.forEach((item) => {
      if (decisions[item.id]) {
        return;
      }
      const initial = defaultInvalidWidthDecision(item);
      onDecisionChange(item.id, initial.action, initial.widthMm);
    });
  }, [decisions, onDecisionChange, invalidLines]);

  if (!needsAttention) {
    return null;
  }

  const allDecided = invalidLines.every((item) => {
    const decision = decisions[item.id] ?? defaultInvalidWidthDecision(item);
    if (decision.action === "exclude") {
      return true;
    }
    return decision.action === "replace_width" && decision.widthMm != null;
  });

  return (
    <Card
      title="Нестандартная ширина"
      subtitle="Завод такую ширину не режет — выберите рез или исключите позицию."
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
          {expanded ? "▾" : "▸"} {invalidLines.length} позиций требуют внимания
        </button>
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      {!expanded ? (
        <p style={{ margin: 0, color: "#475467" }}>
          Разверните блок, чтобы выбрать заводской рез или исключить позицию, затем нажмите «Применить».
        </p>
      ) : (
        <div style={{ display: "grid", gap: "1rem" }}>
          {invalidLines.map((item) => {
            const currentDecision = decisions[item.id] ?? defaultInvalidWidthDecision(item);
            const groupName = `invalid-width-${item.id}`;

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
                  <strong>{item.name || item.line}</strong>
                  <div style={{ color: "#475467", marginTop: "0.25rem" }}>Количество: {item.qty}</div>
                </div>

                <fieldset style={{ border: 0, margin: 0, padding: 0, display: "grid", gap: "0.5rem" }}>
                  <legend style={{ fontSize: "0.9rem", color: "#344054", padding: 0 }}>Действие</legend>
                  {item.replacements.map((repl) => {
                    const value = `width:${repl.width_mm}`;
                    const priceLabel =
                      repl.price != null && repl.price > 0 ? ` — ${formatPrice(repl.price)} ₽` : "";
                    return (
                      <label
                        key={repl.width_mm}
                        style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}
                      >
                        <input
                          type="radio"
                          name={groupName}
                          value={value}
                          checked={choiceValue(currentDecision) === value}
                          onChange={() => onDecisionChange(item.id, "replace_width", repl.width_mm)}
                        />
                        {repl.width_label}
                        {priceLabel}
                      </label>
                    );
                  })}
                  <label style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
                    <input
                      type="radio"
                      name={groupName}
                      value="exclude"
                      checked={currentDecision.action === "exclude"}
                      onChange={() => onDecisionChange(item.id, "exclude", null)}
                    />
                    Исключить позицию
                  </label>
                </fieldset>
              </div>
            );
          })}

          <Button type="button" onClick={onApply} disabled={isPending || !allDecided}>
            {isPending ? "Применяем..." : "Применить"}
          </Button>
        </div>
      )}
    </Card>
  );
};
