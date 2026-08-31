import type { ProductType } from "@/features/commercial-offer/types/commercialOffer";
import { Card } from "@/shared/ui/Card";

type ProductTypePickerProps = {
  onSelect: (productType: ProductType) => void;
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

export const ProductTypePicker = ({ onSelect }: ProductTypePickerProps) => (
  <div style={{ display: "grid", gap: "1.25rem" }}>
    <header>
      <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Создание коммерческого предложения</h1>
      <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
        Выберите тип продукта — в одном КП только один тип: плиты, сваи, ступени, марши, мостовые сваи или ФБС.
      </p>
    </header>

    <div
      style={{
        display: "grid",
        gap: "1rem",
        gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
      }}
    >
      {options.map((option) => (
        <button
          key={option.id}
          type="button"
          onClick={() => onSelect(option.id)}
          style={{
            border: "1px solid #e4e7ec",
            borderRadius: 16,
            padding: "1.25rem",
            background: "#ffffff",
            textAlign: "left",
            cursor: "pointer",
            font: "inherit",
            color: "inherit",
            transition: "border-color 0.15s ease, box-shadow 0.15s ease",
          }}
          onMouseEnter={(event) => {
            event.currentTarget.style.borderColor = "#84adff";
            event.currentTarget.style.boxShadow = "0 8px 24px rgba(43, 92, 255, 0.08)";
          }}
          onMouseLeave={(event) => {
            event.currentTarget.style.borderColor = "#e4e7ec";
            event.currentTarget.style.boxShadow = "none";
          }}
        >
          <Card title={`${option.emoji} ${option.title}`} subtitle={option.description}>
            <div style={{ color: "#175cd3", fontWeight: 600 }}>Выбрать →</div>
          </Card>
        </button>
      ))}
    </div>
  </div>
);
