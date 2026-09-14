import type {
  ArchiveOfferListItem,
  ArchiveSection,
  ProductType,
} from "@/features/commercial-archive/types/archive";
import { formatMoney, truncate } from "@/features/commercial-archive/lib/format";
import {
  holdBadgeLabel,
  holdCreatedByTitle,
  usePromiseHoldsMap,
} from "@/features/factory-capacity/api/promiseQuote";

type Props = {
  section: ArchiveSection;
  items: ArchiveOfferListItem[];
  onSelect: (kpId: number) => void;
  sectionForItem?: (item: ArchiveOfferListItem) => ArchiveSection;
  emptyMessage?: string;
};

const productTypeBadge = (productType: ProductType | undefined): string => {
  if (productType === "piles") {
    return "Сваи";
  }
  if (productType === "steps") {
    return "Ступени";
  }
  if (productType === "marches") {
    return "Марши";
  }
  if (productType === "bridge_piles") {
    return "Мостовые сваи";
  }
  if (productType === "fbs") {
    return "ФБС";
  }
  return "Плиты";
};

const productTypeBadgeStyle = (productType: ProductType | undefined): { background: string; color: string } => {
  if (productType === "piles") {
    return { background: "#ecfdf3", color: "#027a48" };
  }
  if (productType === "steps") {
    return { background: "#fff6ed", color: "#b54708" };
  }
  if (productType === "marches") {
    return { background: "#f4f3ff", color: "#5925dc" };
  }
  if (productType === "bridge_piles") {
    return { background: "#f0f9ff", color: "#026aa2" };
  }
  if (productType === "fbs") {
    return { background: "#eef4ff", color: "#3538cd" };
  }
  return { background: "#eef4ff", color: "#1d4ed8" };
};

const resolveBadgeTypes = (item: ArchiveOfferListItem): ProductType[] => {
  const fromList = (item.product_types ?? []).filter(
    (type): type is ProductType => Boolean(type) && type !== "mixed",
  );
  if (fromList.length > 0) {
    return fromList;
  }
  if (item.product_type && item.product_type !== "mixed") {
    return [item.product_type];
  }
  return ["plates"];
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
  const archivedIds = items
    .filter((item) => (sectionForItem ? sectionForItem(item) : section) === "archived")
    .map((item) => item.kp_id);
  const holds = usePromiseHoldsMap(archivedIds);

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
              <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", flexWrap: "wrap" }}>
                <span style={{ fontWeight: 700 }}>КП №{item.kp_id}</span>
                {resolveBadgeTypes(item).map((type) => (
                  <span
                    key={type}
                    style={{
                      fontSize: "0.75rem",
                      fontWeight: 600,
                      padding: "0.15rem 0.5rem",
                      borderRadius: 999,
                      ...productTypeBadgeStyle(type),
                    }}
                  >
                    {productTypeBadge(type)}
                  </span>
                ))}
                {item.has_delivery_schedule === true && (
                  <span
                    style={{
                      fontSize: "0.75rem",
                      fontWeight: 600,
                      padding: "0.15rem 0.5rem",
                      borderRadius: 999,
                      background: "#ecfdf3",
                      color: "#067647",
                    }}
                  >
                    есть график
                  </span>
                )}
                {holds.get(item.kp_id) && (
                  <span
                    data-testid="promise-hold-badge"
                    title={holdCreatedByTitle(holds.get(item.kp_id)?.created_by)}
                    style={{
                      fontSize: "0.75rem",
                      fontWeight: 600,
                      padding: "0.15rem 0.5rem",
                      borderRadius: 999,
                      background: "#fff6ed",
                      color: "#b54708",
                    }}
                  >
                    {holdBadgeLabel(holds.get(item.kp_id)?.promised_date)}
                  </span>
                )}
              </div>
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
              {item.shipped_progress && item.shipped_progress.x > 0 && (
                <span>
                  отгружено {item.shipped_progress.x}/{item.shipped_progress.m}
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
