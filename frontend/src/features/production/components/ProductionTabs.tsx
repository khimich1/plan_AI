import type { ProductionTab } from "@/features/production/types/production";

type Props = {
  value: ProductionTab;
  onChange: (next: ProductionTab) => void;
};

/** Вкладки UI. `create` остаётся в ProductionTab для programmatic routing с корзины. */
const OPTIONS: { value: ProductionTab; label: string; emoji: string }[] = [
  { value: "calendar", label: "Календарный план", emoji: "📅" },
  { value: "plans", label: "Планы", emoji: "📋" },
  { value: "in-work", label: "КП в работе", emoji: "🧱" },
  { value: "sgp", label: "Склад готовой продукции", emoji: "🏭" },
  { value: "work-calendar", label: "Производственный календарь", emoji: "🗓️" },
];

export const ProductionTabs = ({ value, onChange }: Props) => (
  <div
    role="tablist"
    aria-label="Разделы производства"
    style={{
      display: "inline-flex",
      flexWrap: "wrap",
      padding: 4,
      gap: 4,
      borderRadius: 14,
      background: "#eef2ff",
      border: "1px solid #d6defa",
    }}
  >
    {OPTIONS.map((option) => {
      const isActive = option.value === value;
      return (
        <button
          key={option.value}
          role="tab"
          aria-selected={isActive}
          type="button"
          onClick={() => onChange(option.value)}
          style={{
            border: "none",
            borderRadius: 10,
            padding: "0.55rem 0.9rem",
            cursor: "pointer",
            background: isActive ? "#ffffff" : "transparent",
            color: isActive ? "#23366f" : "#475467",
            fontWeight: 600,
            boxShadow: isActive ? "0 4px 12px rgba(15, 23, 42, 0.08)" : undefined,
          }}
        >
          <span style={{ marginRight: "0.4rem" }} aria-hidden>
            {option.emoji}
          </span>
          {option.label}
        </button>
      );
    })}
  </div>
);
