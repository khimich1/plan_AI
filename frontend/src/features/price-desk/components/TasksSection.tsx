import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import {
  useGuidTasksQuery,
  useResolvePriceMutation,
} from "@/features/nomenclature/hooks/useNomenclatureQueries";
import { productKindLabel } from "@/features/nomenclature/types/nomenclature";
import { getErrorMessage } from "@/shared/lib/apiError";

export const TasksSection = () => {
  const tasksQuery = useGuidTasksQuery();
  const resolvePrice = useResolvePriceMutation();
  const [priceDrafts, setPriceDrafts] = useState<Record<string, string>>({});
  const [localError, setLocalError] = useState<string | null>(null);

  const toCreate = tasksQuery.data?.to_create_1c ?? [];
  const toPrice = tasksQuery.data?.to_price ?? [];
  const empty = toCreate.length === 0 && toPrice.length === 0;

  const submitPrice = async (guid: string) => {
    const raw = priceDrafts[guid] ?? "";
    const normalized = raw.replace(",", ".").trim();
    const price = Number(normalized);
    if (!Number.isFinite(price) || price <= 0) {
      setLocalError("Цена должна быть больше нуля.");
      return;
    }
    setLocalError(null);
    try {
      await resolvePrice.mutateAsync({ guid, price });
      setPriceDrafts((prev) => {
        const next = { ...prev };
        delete next[guid];
        return next;
      });
    } catch (err) {
      setLocalError(getErrorMessage(err));
    }
  };

  return (
    <section aria-label="Задачи по изделиям" style={{ display: "grid", gap: "0.85rem" }}>
      <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Задачи по изделиям</h2>
      {tasksQuery.isError ? (
        <Alert tone="error">{getErrorMessage(tasksQuery.error)}</Alert>
      ) : null}
      {localError ? <Alert tone="error">{localError}</Alert> : null}
      {empty && !tasksQuery.isLoading ? (
        <p style={{ margin: 0, color: "#667085" }}>задач нет</p>
      ) : null}

      {toCreate.length > 0 ? (
        <div>
          <h3 style={{ margin: "0 0 0.45rem", fontSize: "0.95rem" }}>Завести в 1С</h3>
          <ul style={{ margin: 0, paddingLeft: "1.15rem", display: "grid", gap: "0.35rem" }}>
            {toCreate.map((item) => (
              <li key={`${item.product_kind}-${item.mark}-${item.field ?? "guid_1c"}`}>
                {productKindLabel(item.product_kind)} · {item.mark} — {item.hint}
              </li>
            ))}
          </ul>
        </div>
      ) : null}

      {toPrice.length > 0 ? (
        <div>
          <h3 style={{ margin: "0 0 0.45rem", fontSize: "0.95rem" }}>Ввести цену</h3>
          <ul style={{ listStyle: "none", margin: 0, padding: 0, display: "grid", gap: "0.65rem" }}>
            {toPrice.map((item) => (
              <li
                key={item.guid}
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  gap: "0.5rem",
                  alignItems: "center",
                }}
              >
                <span>
                  {productKindLabel(item.product_kind)} · {item.mark}
                </span>
                <input
                  aria-label={`Цена для ${item.mark}`}
                  value={priceDrafts[item.guid] ?? ""}
                  onChange={(event) =>
                    setPriceDrafts((prev) => ({ ...prev, [item.guid]: event.target.value }))
                  }
                  inputMode="decimal"
                  style={{
                    width: 120,
                    border: "1px solid #d0d5dd",
                    borderRadius: 8,
                    padding: "0.4rem 0.55rem",
                  }}
                />
                <Button
                  type="button"
                  onClick={() => void submitPrice(item.guid)}
                  disabled={resolvePrice.isPending}
                >
                  Записать цену
                </Button>
              </li>
            ))}
          </ul>
        </div>
      ) : null}
    </section>
  );
};
