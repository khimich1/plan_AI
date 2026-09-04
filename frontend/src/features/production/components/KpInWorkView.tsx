import { useState, type CSSProperties } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Card } from "@/shared/ui/Card";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import { useKpCandidatesQuery } from "@/features/production/hooks/useProductionQueries";
import type {
  KpCandidateItem,
  KpCandidatePlateItem,
} from "@/features/production/types/production";

const EMPTY_MESSAGE = "Все плиты на СГП — смотрите склад готовой продукции";

const rowStyle: CSSProperties = {
  display: "grid",
  gap: "0.45rem",
  padding: "0.9rem 1rem",
  border: "1px solid #e4e7ec",
  borderRadius: 14,
  background: "#ffffff",
  cursor: "pointer",
  textAlign: "left",
  width: "100%",
  font: "inherit",
  color: "inherit",
};

const formatTerms = (value: string): string => {
  const trimmed = value.trim();
  return trimmed ? trimmed : "без срока";
};

const leftoverQty = (item: KpCandidateItem): number =>
  (item.remaining_qty ?? 0) + (item.in_plan_qty ?? 0);

const formatDims = (plate: KpCandidatePlateItem): string => {
  const length = Number.isFinite(plate.length_m) ? plate.length_m.toFixed(2) : "—";
  const width = Number.isFinite(plate.width_m) ? plate.width_m.toFixed(2) : "—";
  return `${length} × ${width} м`;
};

const bucketLabel = (bucket: KpCandidatePlateItem["bucket"]): string =>
  bucket === "in_plan" ? "в плане, ждёт отливки" : "ждёт плана";

export const KpInWorkView = () => {
  const query = useKpCandidatesQuery(true, "in_work");
  const [openKpId, setOpenKpId] = useState<number | null>(null);

  if (query.isLoading) {
    return (
      <Card title="КП в работе">
        <div style={{ display: "flex", gap: "0.5rem" }}>
          <Spinner /> Загрузка…
        </div>
      </Card>
    );
  }

  if (query.isError) {
    return (
      <Card title="КП в работе">
        <Alert tone="error">{getErrorMessage(query.error)}</Alert>
      </Card>
    );
  }

  const items = query.data?.items ?? [];

  if (items.length === 0) {
    return (
      <Card title="КП в работе">
        <Alert tone="info">{EMPTY_MESSAGE}</Alert>
      </Card>
    );
  }

  return (
    <Card
      title="КП в работе"
      subtitle="Очередь плит, которые ещё не на складе готовой продукции"
    >
      <div style={{ display: "grid", gap: "0.6rem" }}>
        {items.map((item) => {
          const isOpen = openKpId === item.kp_id;
          const label = `КП №${item.kp_id}`;
          return (
            <div key={item.kp_id}>
              <button
                type="button"
                onClick={() => setOpenKpId(isOpen ? null : item.kp_id)}
                aria-expanded={isOpen}
                aria-label={isOpen ? `Свернуть ${label}` : `Развернуть ${label}`}
                style={rowStyle}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    gap: "0.75rem",
                    flexWrap: "wrap",
                    alignItems: "baseline",
                  }}
                >
                  <span style={{ fontWeight: 700 }}>
                    {isOpen ? "▾" : "▸"} {label}
                  </span>
                  <span style={{ color: "#475467" }}>{formatTerms(item.execution_terms)}</span>
                </div>
                <div style={{ color: "#475467", display: "flex", gap: "0.85rem", flexWrap: "wrap" }}>
                  <span>{item.customer_name || "Без клиента"}</span>
                  <span>осталось {leftoverQty(item)} шт</span>
                  <span>в плане {item.in_plan_qty ?? 0}</span>
                  <span>на СГП {item.on_sgp_qty ?? 0}</span>
                </div>
              </button>
              {isOpen && <PlateList plates={item.plates} />}
            </div>
          );
        })}
      </div>
    </Card>
  );
};

const PlateList = ({ plates }: { plates: KpCandidatePlateItem[] }) => {
  if (plates.length === 0) {
    return (
      <div style={{ padding: "0.75rem 1rem", color: "#667085" }}>Нет плит не на СГП.</div>
    );
  }

  return (
    <table
      style={{
        width: "100%",
        borderCollapse: "collapse",
        fontSize: "0.9rem",
        marginTop: 4,
      }}
    >
      <thead>
        <tr style={{ textAlign: "left", color: "#667085" }}>
          <th style={{ padding: "0.45rem 0.75rem" }}>Марка</th>
          <th style={{ padding: "0.45rem 0.75rem" }}>Размер</th>
          <th style={{ padding: "0.45rem 0.75rem" }}>Нагрузка</th>
          <th style={{ padding: "0.45rem 0.75rem" }}>Шт</th>
          <th style={{ padding: "0.45rem 0.75rem" }}>Статус</th>
        </tr>
      </thead>
      <tbody>
        {plates.map((plate) => (
          <tr key={plate.id} style={{ borderTop: "1px solid #eef2f6" }}>
            <td style={{ padding: "0.45rem 0.75rem", fontWeight: 600 }}>{plate.plate_name}</td>
            <td style={{ padding: "0.45rem 0.75rem" }}>{formatDims(plate)}</td>
            <td style={{ padding: "0.45rem 0.75rem" }}>
              {plate.load_class != null ? plate.load_class : "—"}
            </td>
            <td style={{ padding: "0.45rem 0.75rem" }}>{plate.qty}</td>
            <td style={{ padding: "0.45rem 0.75rem", color: "#475467" }}>
              {bucketLabel(plate.bucket)}
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
};
