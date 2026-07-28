import type {
  ArchiveOfferListItem,
  ArchiveSection,
} from "@/features/commercial-archive/types/archive";
import { formatMoney, truncate } from "@/features/commercial-archive/lib/format";

type Props = {
  section: ArchiveSection;
  items: ArchiveOfferListItem[];
  onSelect: (kpId: number) => void;
  sectionForItem?: (item: ArchiveOfferListItem) => ArchiveSection;
  emptyMessage?: string;
};

const rowStyle: React.CSSProperties = {
  display: "grid",
  gap: "0.5rem",
  padding: "0.9rem 1rem",
  border: "1px solid #e4e7ec",
  borderRadius: 14,
  background: "#ffffff",
  cursor: "pointer",
  textAlign: "left",
  width: "100%",
  font: "inherit",
  color: "inherit",
  transition: "border-color 0.15s ease, box-shadow 0.15s ease",
};

const trailingMeta = (section: ArchiveSection, item: ArchiveOfferListItem): string => {
  switch (section) {
    case "archived":
      return `${item.discount_percent}%`;
    case "in_production":
      return item.execution_terms ? `⏰ ${item.execution_terms}` : "без срока";
    case "completed":
      return item.creation_date ? `📅 ${item.creation_date}` : "";
    default:
      return "";
  }
};

export const ArchiveOfferList = ({
  section,
  items,
  onSelect,
  sectionForItem,
  emptyMessage = "В этом разделе пока нет КП.",
}: Props) => {
  if (items.length === 0) {
    return (
      <div
        style={{
          padding: "2rem",
          border: "1px dashed #d0d5dd",
          borderRadius: 16,
          background: "#ffffff",
          color: "#667085",
          textAlign: "center",
        }}
      >
        {emptyMessage}
      </div>
    );
  }

  return (
    <div style={{ display: "grid", gap: "0.6rem" }}>
      {items.map((item) => {
        const itemSection = sectionForItem ? sectionForItem(item) : section;
        const trailing = trailingMeta(itemSection, item);
        const percentBadge =
          (itemSection === "in_production" || itemSection === "completed") &&
          item.completion_percentage !== null
            ? `${item.completion_percentage.toFixed(1)}%`
            : null;

        return (
          <button
            key={item.kp_id}
            type="button"
            onClick={() => onSelect(item.kp_id)}
            style={rowStyle}
            onMouseEnter={(event) => {
              event.currentTarget.style.borderColor = "#b4bfff";
              event.currentTarget.style.boxShadow = "0 6px 16px rgba(15, 23, 42, 0.06)";
            }}
            onMouseLeave={(event) => {
              event.currentTarget.style.borderColor = "#e4e7ec";
              event.currentTarget.style.boxShadow = "none";
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem", flexWrap: "wrap" }}>
              <div style={{ fontWeight: 700 }}>КП №{item.kp_id}</div>
              <div style={{ color: "#101828", fontWeight: 600 }}>{formatMoney(item.total_amount)}</div>
            </div>
            <div style={{ color: "#475467", display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
              <span>{truncate(item.customer_name || "Без клиента", 32)}</span>
              {item.manager_name && (
                <span style={{ color: "#667085" }}>· {truncate(item.manager_name, 22)}</span>
              )}
            </div>
            <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap", fontSize: "0.9rem", color: "#667085" }}>
              {item.status === "На СГП" && <span>🏬 На СГП</span>}
              {item.sgp_progress && item.sgp_progress.m > 0 && (
                <span>
                  {item.sgp_progress.n}/{item.sgp_progress.m} на СГП
                </span>
              )}
              {percentBadge && <span>Готовность: {percentBadge}</span>}
              {trailing && <span>{trailing}</span>}
            </div>
          </button>
        );
      })}
    </div>
  );
};
