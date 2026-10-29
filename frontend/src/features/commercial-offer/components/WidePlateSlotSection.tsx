import { useState } from "react";

import { uniqueWideSlotsByLine, type WideSlot, type WideSlotAction } from "@/features/commercial-offer/lib/widePlateSlots";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Textarea } from "@/shared/ui/Field";

type WidePlateSlotSectionProps = {
  lines: readonly string[];
  slots: readonly WideSlot[];
  onAccept?: (slotId: string) => void;
  onApply?: (slotIds: string[]) => void;
  onAction?: (slotId: string, action: WideSlotAction) => void;
  onReplacement?: (slotId: string, replacementText: string) => void;
  errorMessage?: string | null;
  /** Photo page id or text-review. A new key collapses the list again. */
  collapseKey?: string;
};

const SLOT_ACTIONS: Array<[WideSlotAction, string]> = [
  ["confirm", "Подтвердить"],
  ["replace", "Заменить"],
  ["exclude", "Исключить"],
];

export const WidePlateSlotSection = ({
  lines,
  slots,
  onAccept,
  onApply,
  onAction,
  onReplacement,
  errorMessage,
  collapseKey,
}: WidePlateSlotSectionProps) => {
  const openSlots = uniqueWideSlotsByLine(slots).filter((slot) => !slot.accepted);
  const [expanded, setExpanded] = useState(false);
  const [seenCollapseKey, setSeenCollapseKey] = useState(collapseKey);
  if (collapseKey !== seenCollapseKey) {
    setSeenCollapseKey(collapseKey);
    setExpanded(false);
  }
  if (openSlots.length === 0) {
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
          {expanded ? "▾" : "▸"} {openSlots.length} позиций требуют внимания
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
          {openSlots.map((slot) => (
            <div
              key={slot.id}
              style={{
                display: "grid",
                gap: "0.75rem",
                border: "1px solid #e4e7ec",
                borderRadius: 12,
                padding: "1rem",
                background: "#fafafa",
              }}
            >
              <strong>{lines[slot.lineIndex] ?? ""}</strong>
              <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap", alignItems: "center" }}>
                {SLOT_ACTIONS.map(([value, label]) => (
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
                      name={`wide-slot-${slot.id}`}
                      value={value}
                      checked={slot.action === value}
                      onChange={() => onAction?.(slot.id, value)}
                    />
                    <span>{label}</span>
                  </label>
                ))}
                <Button type="button" variant="secondary" onClick={() => onAccept?.(slot.id)}>
                  ОК
                </Button>
              </div>
              {slot.action === "replace" && (
                <FieldWrapper label="Заменяющие позиции" hint="Каждая строка — отдельная позиция в списке.">
                  <Textarea
                    value={slot.replacementText}
                    onChange={(event) => onReplacement?.(slot.id, event.target.value)}
                  />
                </FieldWrapper>
              )}
            </div>
          ))}
          <Button type="button" onClick={() => onApply?.(openSlots.map((slot) => slot.id))}>
            Применить решения
          </Button>
        </div>
      )}
    </Card>
  );
};
