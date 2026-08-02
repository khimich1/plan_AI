import { useState } from "react";
import { formatWeightKg } from "@/features/logistics/lib/logisticsFormat";
import type { LayoutMetadata, LayoutStack } from "@/features/logistics/types/logistics";

type Props = {
  layout: LayoutMetadata | null;
  totalWeightKg?: number | null;
  maxWeightKg?: number | null;
};

const formatMeters = (value: number): string =>
  value.toLocaleString("ru-RU", { maximumFractionDigits: 2 });

const stackWord = (count: number): string => {
  const mod10 = count % 10;
  const mod100 = count % 100;
  if (mod10 === 1 && mod100 !== 11) return "штабель";
  if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) return "штабеля";
  return "штабелей";
};

const STACK_BACKGROUNDS = ["#f2f4f7", "#e4e7ec"];

const cardStyle: React.CSSProperties = {
  border: "1px solid #eaecf0",
  borderRadius: 14,
  background: "#ffffff",
  padding: "1rem",
  display: "grid",
  gap: "0.75rem",
};

const stackButtonStyle = (expanded: boolean, colorIndex: number): React.CSSProperties => ({
  border: expanded ? "2px solid #475467" : "1px solid #d0d5dd",
  borderRadius: 10,
  background: STACK_BACKGROUNDS[colorIndex % STACK_BACKGROUNDS.length],
  padding: "0.6rem 0.4rem",
  cursor: "pointer",
  fontSize: "0.85rem",
  fontWeight: 600,
  color: "#1d2939",
  textAlign: "center",
  minWidth: 0,
});

const tierUnitText = (stack: LayoutStack, tierIndex: number): string => {
  const tier = stack.tiers.find((t) => t.index === tierIndex);
  if (!tier) return "";
  return tier.units
    .map((unit) =>
      unit.width_m != null
        ? `${unit.plate_name} (${formatMeters(unit.width_m)} м)`
        : unit.plate_name,
    )
    .join(" + ");
};

export const LayoutBlock = ({ layout, totalWeightKg, maxWeightKg }: Props) => {
  const [expandedIndex, setExpandedIndex] = useState<number | null>(null);

  if (!layout || layout.stacks.length === 0) {
    return null;
  }

  const expandedStack = layout.stacks.find((stack) => stack.index === expandedIndex) ?? null;
  const weightPart =
    totalWeightKg != null
      ? ` · ${formatWeightKg(totalWeightKg)}${maxWeightKg != null ? ` / ${formatWeightKg(maxWeightKg)}` : ""}`
      : "";

  return (
    <section style={cardStyle} aria-label="Укладка в кузов">
      <div style={{ fontWeight: 600, fontSize: "0.95rem", color: "#1d2939" }}>
        Укладка в кузов · {layout.stacks.length} {stackWord(layout.stacks.length)} ·{" "}
        {formatMeters(layout.body_used_m)} / {formatMeters(layout.body_length_m)} м{weightPart}
      </div>

      <div style={{ display: "flex", gap: "0.4rem", width: "100%", overflowX: "auto" }}>
        {layout.stacks.map((stack, colorIndex) => {
          const share = (stack.marking_length_m / layout.body_length_m) * 100;
          const isExpanded = stack.index === expandedIndex;
          return (
            <button
              key={stack.index}
              type="button"
              aria-expanded={isExpanded}
              onClick={() => setExpandedIndex(isExpanded ? null : stack.index)}
              style={{
                ...stackButtonStyle(isExpanded, colorIndex),
                flex: `0 0 calc(${share}% - 0.4rem)`,
              }}
            >
              Штабель {stack.index}
              <div style={{ fontWeight: 400, color: "#475467" }}>
                {formatMeters(stack.marking_length_m)} м · {stack.tiers.length} яр.
              </div>
            </button>
          );
        })}
      </div>

      {expandedStack && (
        <div style={{ display: "grid", gap: "0.3rem", fontSize: "0.9rem" }}>
          {[...expandedStack.tiers]
            .sort((a, b) => b.index - a.index)
            .map((tier) => (
              <div key={tier.index}>
                <span style={{ color: "#475467" }}>Ярус {tier.index}:</span>{" "}
                {tierUnitText(expandedStack, tier.index)}
              </div>
            ))}
        </div>
      )}

      {layout.loading_steps.length > 0 && (
        <div>
          <div style={{ fontWeight: 600, fontSize: "0.9rem", marginBottom: "0.3rem" }}>
            Порядок погрузки
          </div>
          <ol style={{ margin: 0, paddingLeft: "1.2rem", display: "grid", gap: "0.2rem", fontSize: "0.9rem" }}>
            {layout.loading_steps.map((step) => (
              <li key={step.step}>
                Штабель {step.stack_index}, ярус {step.tier_index}: {step.description}
              </li>
            ))}
          </ol>
        </div>
      )}
    </section>
  );
};
