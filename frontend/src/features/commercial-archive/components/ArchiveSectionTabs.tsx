import type { ArchiveSection } from "@/features/commercial-archive/types/archive";

type Props = {
  value: ArchiveSection;
  onChange: (section: ArchiveSection) => void;
};

const OPTIONS: { value: ArchiveSection; label: string; emoji: string }[] = [
  { value: "archived", label: "В архиве", emoji: "📦" },
  { value: "in_production", label: "В производстве", emoji: "🏭" },
  { value: "completed", label: "Выполненные", emoji: "✅" },
];

export const ArchiveSectionTabs = ({ value, onChange }: Props) => (
  <div
    role="tablist"
    aria-label="Разделы архива"
    style={{
      display: "inline-flex",
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
