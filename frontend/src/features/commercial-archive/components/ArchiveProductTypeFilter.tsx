import type { ArchiveProductTypeFilter } from "@/features/commercial-archive/types/archive";

type Props = {
  value: ArchiveProductTypeFilter;
  onChange: (filter: ArchiveProductTypeFilter) => void;
};

const OPTIONS: { value: ArchiveProductTypeFilter; label: string }[] = [
  { value: "all", label: "Все" },
  { value: "plates", label: "Плиты" },
  { value: "piles", label: "Сваи" },
  { value: "steps", label: "Ступени" },
  { value: "marches", label: "Марши" },
  { value: "bridge_piles", label: "Мостовые сваи" },
  { value: "fbs", label: "ФБС" },
];

export const ArchiveProductTypeFilterTabs = ({ value, onChange }: Props) => (
  <div
    role="tablist"
    aria-label="Тип продукции"
    style={{
      display: "inline-flex",
      padding: 4,
      gap: 4,
      borderRadius: 14,
      background: "#f2f4f7",
      border: "1px solid #e4e7ec",
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
            padding: "0.45rem 0.85rem",
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
