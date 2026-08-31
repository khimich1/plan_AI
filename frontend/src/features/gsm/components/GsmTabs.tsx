import type { GsmTab } from "@/features/gsm/types/gsm";

type Props = {
  value: GsmTab;
  onChange: (next: GsmTab) => void;
};

const OPTIONS: { value: GsmTab; label: string }[] = [
  { value: "overview", label: "Обзор" },
  { value: "transactions", label: "Транзакции" },
  { value: "registries", label: "Справочники" },
];

export const GsmTabs = ({ value, onChange }: Props) => (
  <div
    role="tablist"
    aria-label="Разделы ГСМ"
    style={{
      display: "inline-flex",
      flexWrap: "wrap",
      padding: 4,
      gap: 4,
      borderRadius: 14,
      background: "#eef2ff",
      border: "1px solid #d6defa",
      alignSelf: "flex-start",
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
          {option.label}
        </button>
      );
    })}
  </div>
);
