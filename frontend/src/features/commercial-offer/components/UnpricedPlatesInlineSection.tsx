import { useEffect, useState } from "react";
import type {
  CommercialDraftDetails,
  UnpricedPlateAction,
  UnpricedPlateLine,
} from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

export type UnpricedPlateDecisionState = {
  action: UnpricedPlateAction;
  loadCode: number | null;
};

type UnpricedPlatesInlineSectionProps = {
  draft: CommercialDraftDetails;
  decisions: Record<string, UnpricedPlateDecisionState>;
  isPending: boolean;
  errorMessage?: string | null;
  onDecisionChange: (lineId: string, action: UnpricedPlateAction, loadCode: number | null) => void;
  onApply: () => void;
};

const formatPrice = (price: number) =>
  new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 }).format(price);

const defaultDecisionForLine = (item: UnpricedPlateLine): UnpricedPlateDecisionState => {
  const first = item.replacements[0];
  if (first) {
    return { action: "replace_load", loadCode: first.load_code };
  }
  return { action: "exclude", loadCode: null };
};

const choiceValue = (decision: UnpricedPlateDecisionState): string => {
  if (decision.action === "exclude") {
    return "exclude";
  }
  return `load:${decision.loadCode ?? ""}`;
};

export const UnpricedPlatesInlineSection = ({
  draft,
  decisions,
  isPending,
  errorMessage,
  onDecisionChange,
  onApply,
}: UnpricedPlatesInlineSectionProps) => {
  const unpricedLines = draft.metadata.unpriced_plate_lines ?? [];
  const [expanded, setExpanded] = useState(false);
  const needsAttention = unpricedLines.length > 0 && !draft.metadata.unpriced_plates_resolved;

  useEffect(() => {
    unpricedLines.forEach((item) => {
      if (decisions[item.id]) {
        return;
      }
      const initial = defaultDecisionForLine(item);
      onDecisionChange(item.id, initial.action, initial.loadCode);
    });
  }, [decisions, onDecisionChange, unpricedLines]);

  if (!needsAttention) {
    return null;
  }

  const allDecided = unpricedLines.every((item) => {
    const decision = decisions[item.id] ?? defaultDecisionForLine(item);
    if (decision.action === "exclude") {
      return true;
    }
    return decision.action === "replace_load" && decision.loadCode != null;
  });

  return (
    <Card
      title="Нет в прайсе / не производится"
      subtitle="Замена нагрузки требует согласования с заказчиком."
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
          {expanded ? "▾" : "▸"} {unpricedLines.length} позиций требуют внимания
        </button>
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      {!expanded ? (
        <p style={{ margin: 0, color: "#475467" }}>
          Разверните блок, чтобы выбрать замену нагрузки или исключить позицию, затем нажмите «Применить».
        </p>
      ) : (
        <div style={{ display: "grid", gap: "1rem" }}>
          {unpricedLines.map((item) => {
            const currentDecision = decisions[item.id] ?? defaultDecisionForLine(item);
            const hasReplacements = item.replacements.length > 0;

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
                  {!hasReplacements && (
                    <div style={{ color: "#b54708", marginTop: "0.35rem", fontSize: "0.9rem" }}>
                      Производимых замен нет — доступно только исключение позиции.
                    </div>
                  )}
                </div>

                <label style={{ display: "grid", gap: "0.35rem" }}>
                  <span style={{ fontSize: "0.9rem", color: "#344054" }}>Действие</span>
                  <select
                    value={choiceValue(currentDecision)}
                    onChange={(event) => {
                      const value = event.target.value;
                      if (value === "exclude") {
                        onDecisionChange(item.id, "exclude", null);
                        return;
                      }
                      const loadCode = Number(value.replace("load:", ""));
                      onDecisionChange(item.id, "replace_load", Number.isFinite(loadCode) ? loadCode : null);
                    }}
                    style={{
                      border: "1px solid #d0d5dd",
                      borderRadius: 10,
                      padding: "0.65rem 0.75rem",
                      background: "#ffffff",
                    }}
                  >
                    {item.replacements.map((repl) => (
                      <option key={repl.load_code} value={`load:${repl.load_code}`}>
                        {repl.load_code}п — {formatPrice(repl.price)} ₽
                      </option>
                    ))}
                    <option value="exclude">Исключить позицию</option>
                  </select>
                </label>
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
