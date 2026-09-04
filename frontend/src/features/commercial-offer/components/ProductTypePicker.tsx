import { useMemo, useState } from "react";
import type { CSSProperties } from "react";
import type { ProductType } from "@/features/commercial-offer/types/commercialOffer";
import { Card } from "@/shared/ui/Card";
import { Drawer } from "@/shared/ui/Drawer";

type ProductTypePickerProps = {
  mode?: "create" | "append";
  selectedProductTypes?: ReadonlyArray<ProductType>;
  orderLines?: ReadonlyArray<Record<string, unknown>>;
  managerName?: string;
  clientName?: string;
  onSelect: (productType: ProductType) => void;
  onBackToResult?: () => void;
};

const PRODUCT_TYPE_LABELS: Record<ProductType, string> = {
  plates: "Плиты",
  piles: "Сваи",
  steps: "Ступени",
  marches: "Марши",
  bridge_piles: "Мостовые сваи",
  fbs: "ФБС",
};

const options: Array<{ id: ProductType; title: string; description: string; emoji: string }> = [
  {
    id: "plates",
    title: "Плиты",
    description: "Коммерческое предложение на железобетонные плиты перекрытия.",
    emoji: "🧱",
  },
  {
    id: "piles",
    title: "Сваи",
    description: "Коммерческое предложение на цельные железобетонные сваи.",
    emoji: "🏗️",
  },
  {
    id: "steps",
    title: "Ступени",
    description: "Коммерческое предложение на лестничные ступени.",
    emoji: "🪜",
  },
  {
    id: "marches",
    title: "Марши",
    description: "Коммерческое предложение на лестничные марши.",
    emoji: "🪜",
  },
  {
    id: "bridge_piles",
    title: "Мостовые сваи",
    description: "Коммерческое предложение на мостовые железобетонные сваи.",
    emoji: "🌉",
  },
  {
    id: "fbs",
    title: "ФБС",
    description: "Коммерческое предложение на фундаментные блоки ФБС.",
    emoji: "📦",
  },
];

const tileBaseStyle: CSSProperties = {
  border: "1px solid #e4e7ec",
  borderRadius: 16,
  padding: "1.25rem",
  background: "#ffffff",
  textAlign: "left",
  font: "inherit",
  color: "inherit",
  transition: "border-color 0.15s ease, box-shadow 0.15s ease",
  position: "relative",
};

const applyHover = (element: HTMLElement, active: boolean) => {
  element.style.borderColor = active ? "#84adff" : "#e4e7ec";
  element.style.boxShadow = active ? "0 8px 24px rgba(43, 92, 255, 0.08)" : "none";
};

const iconButtonStyle: CSSProperties = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  minWidth: 40,
  minHeight: 40,
  width: 44,
  height: 44,
  border: "1px solid #d0d5dd",
  borderRadius: 12,
  background: "#ffffff",
  cursor: "pointer",
  color: "#175cd3",
  padding: 0,
};

const PlusIcon = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" />
  </svg>
);

const InfoIcon = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="2" />
    <path d="M12 10v6" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" />
    <circle cx="12" cy="7" r="1.25" fill="currentColor" />
  </svg>
);

const CheckIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path
      d="M5 12.5 10 17.5 19 7.5"
      stroke="currentColor"
      strokeWidth="2.5"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const lineDisplayName = (line: Record<string, unknown>): string =>
  String(line.name ?? line.mark ?? "").trim() || "—";

const lineQty = (line: Record<string, unknown>): string => {
  const qty = line.qty;
  if (qty == null || qty === "") {
    return "—";
  }
  return String(qty);
};

export const ProductTypePicker = ({
  mode = "create",
  selectedProductTypes = [],
  orderLines = [],
  managerName,
  clientName,
  onSelect,
  onBackToResult,
}: ProductTypePickerProps) => {
  const isAppend = mode === "append";
  const selectedSet = useMemo(() => new Set(selectedProductTypes), [selectedProductTypes]);
  const [infoType, setInfoType] = useState<ProductType | null>(null);

  const stripLabels = selectedProductTypes
    .map((type) => PRODUCT_TYPE_LABELS[type] ?? type)
    .filter(Boolean);

  const drawerLines = useMemo(() => {
    if (!infoType) {
      return [];
    }
    return orderLines.filter((line) => String(line.product_type ?? "") === infoType);
  }, [infoType, orderLines]);

  const drawerTitle = infoType
    ? `Уже в КП · ${PRODUCT_TYPE_LABELS[infoType] ?? infoType}`
    : "Уже в КП";

  return (
    <div style={{ display: "grid", gap: "1.25rem" }}>
      <header style={{ display: "grid", gap: "0.75rem" }}>
        {isAppend ? (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", gap: "1rem", flexWrap: "wrap" }}>
              <div>
                <h1 style={{ margin: 0, fontSize: "1.75rem" }}>
                  {managerName?.trim() || "Менеджер не указан"}
                </h1>
                <p style={{ margin: "0.35rem 0 0", color: "#344054", fontSize: "1.05rem" }}>
                  Заказчик: {clientName?.trim() || "не указан"}
                </p>
              </div>
              {onBackToResult ? (
                <button
                  type="button"
                  onClick={onBackToResult}
                  style={{
                    alignSelf: "flex-start",
                    border: "1px solid #d0d5dd",
                    borderRadius: 12,
                    padding: "0.65rem 1rem",
                    background: "#ffffff",
                    cursor: "pointer",
                    font: "inherit",
                    fontWeight: 600,
                    color: "#344054",
                  }}
                >
                  К результату
                </button>
              ) : null}
            </div>
            <p style={{ margin: 0, color: "#475467" }}>
              Выберите тип продукции для дополнения текущего КП.
            </p>
            {stripLabels.length > 0 ? (
              <div
                style={{
                  border: "1px solid #d0d5dd",
                  borderRadius: 12,
                  background: "#f8fafc",
                  padding: "0.75rem 1rem",
                  color: "#344054",
                  fontWeight: 600,
                }}
              >
                Уже в КП: {stripLabels.join(" · ")}
              </div>
            ) : null}
          </>
        ) : (
          <>
            <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Создание коммерческого предложения</h1>
            <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
              Выберите тип продукции: плиты, сваи, ступени, марши, мостовые сваи или ФБС.
            </p>
          </>
        )}
      </header>

      <div
        style={{
          display: "grid",
          gap: "1rem",
          gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
        }}
      >
        {options.map((option) => {
          const isSelected = isAppend && selectedSet.has(option.id);
          const labelLower = PRODUCT_TYPE_LABELS[option.id].toLowerCase();

          if (isSelected) {
            return (
              <div
                key={option.id}
                data-product-type={option.id}
                style={{ ...tileBaseStyle, cursor: "default" }}
                onMouseEnter={(event) => applyHover(event.currentTarget, true)}
                onMouseLeave={(event) => applyHover(event.currentTarget, false)}
              >
                <span
                  aria-label={`Уже добавлен: ${labelLower}`}
                  style={{
                    position: "absolute",
                    top: 12,
                    right: 12,
                    display: "inline-flex",
                    alignItems: "center",
                    justifyContent: "center",
                    width: 28,
                    height: 28,
                    borderRadius: "50%",
                    background: "#ecfdf3",
                    color: "#027a48",
                  }}
                >
                  <CheckIcon />
                </span>
                <Card title={`${option.emoji} ${option.title}`} subtitle={option.description}>
                  <div style={{ display: "flex", gap: "0.75rem", alignItems: "center" }}>
                    <button
                      type="button"
                      aria-label={`Добавить ${labelLower} в КП`}
                      onClick={() => onSelect(option.id)}
                      style={iconButtonStyle}
                    >
                      <PlusIcon />
                    </button>
                    <button
                      type="button"
                      aria-label={`Показать позиции: ${labelLower}`}
                      onClick={() => setInfoType(option.id)}
                      style={iconButtonStyle}
                    >
                      <InfoIcon />
                    </button>
                  </div>
                </Card>
              </div>
            );
          }

          return (
            <button
              key={option.id}
              type="button"
              data-product-type={option.id}
              onClick={() => onSelect(option.id)}
              style={{ ...tileBaseStyle, cursor: "pointer" }}
              onMouseEnter={(event) => applyHover(event.currentTarget, true)}
              onMouseLeave={(event) => applyHover(event.currentTarget, false)}
            >
              <Card title={`${option.emoji} ${option.title}`} subtitle={option.description}>
                <div style={{ color: "#175cd3", fontWeight: 600 }}>Выбрать →</div>
              </Card>
            </button>
          );
        })}
      </div>

      <Drawer
        open={infoType !== null}
        onClose={() => setInfoType(null)}
        title={drawerTitle}
        width={420}
        side="left"
      >
        {drawerLines.length === 0 ? (
          <p style={{ margin: 0, color: "#667085" }}>Нет позиций этого типа.</p>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th
                  style={{
                    textAlign: "left",
                    padding: "0.65rem 0.5rem",
                    borderBottom: "1px solid #e4e7ec",
                    color: "#475467",
                    fontWeight: 600,
                    width: 48,
                  }}
                >
                  №
                </th>
                <th
                  style={{
                    textAlign: "left",
                    padding: "0.65rem 0.5rem",
                    borderBottom: "1px solid #e4e7ec",
                    color: "#475467",
                    fontWeight: 600,
                  }}
                >
                  Наименование
                </th>
                <th
                  style={{
                    textAlign: "right",
                    padding: "0.65rem 0.5rem",
                    borderBottom: "1px solid #e4e7ec",
                    color: "#475467",
                    fontWeight: 600,
                    width: 96,
                  }}
                >
                  Кол-во
                </th>
              </tr>
            </thead>
            <tbody>
              {drawerLines.map((line, index) => (
                <tr key={`${lineDisplayName(line)}-${index}`}>
                  <td
                    style={{
                      padding: "0.7rem 0.5rem",
                      borderBottom: "1px solid #f2f4f7",
                      fontVariantNumeric: "tabular-nums",
                      color: "#667085",
                    }}
                  >
                    {index + 1}
                  </td>
                  <td style={{ padding: "0.7rem 0.5rem", borderBottom: "1px solid #f2f4f7" }}>
                    {lineDisplayName(line)}
                  </td>
                  <td
                    style={{
                      padding: "0.7rem 0.5rem",
                      borderBottom: "1px solid #f2f4f7",
                      textAlign: "right",
                      fontVariantNumeric: "tabular-nums",
                    }}
                  >
                    {lineQty(line)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Drawer>
    </div>
  );
};
