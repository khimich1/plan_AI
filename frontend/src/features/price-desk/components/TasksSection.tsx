import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { GuidExplainPanel } from "@/features/nomenclature/components/GuidExplainPanel";
import { economistGuidExplainSource, useRefreshGuidTasks } from "@/features/nomenclature/hooks/useGuidExplainQueries";
import {
  useGuidTasksQuery,
  useResolvePriceMutation,
} from "@/features/nomenclature/hooks/useNomenclatureQueries";
import { formatKpIdsLabel, productKindLabel } from "@/features/nomenclature/types/nomenclature";
import { getErrorMessage } from "@/shared/lib/apiError";

const canOpenImport = (productKind: string): boolean =>
  productKind !== "plate" && productKind !== "plates";

type TasksSectionProps = {
  onOpenImport: () => void;
};

export const TasksSection = ({ onOpenImport }: TasksSectionProps) => {
  const tasksQuery = useGuidTasksQuery();
  const resolvePrice = useResolvePriceMutation();
  const refreshTasks = useRefreshGuidTasks();
  const [priceDrafts, setPriceDrafts] = useState<Record<string, string>>({});
  const [localError, setLocalError] = useState<string | null>(null);

  const toCreate = tasksQuery.data?.to_create_1c ?? [];
  const toPrice = tasksQuery.data?.to_price ?? [];

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

  if (!tasksQuery.isError && !localError && toCreate.length === 0 && toPrice.length === 0) {
    return null;
  }

  return (
    <section aria-label="Задачи по изделиям" style={{ display: "grid", gap: "0.85rem" }}>
      {tasksQuery.isError ? (
        <Alert tone="error">{getErrorMessage(tasksQuery.error)}</Alert>
      ) : null}
      {localError ? <Alert tone="error">{localError}</Alert> : null}

      {toCreate.length > 0 ? (
        <div>
          <h3 style={{ margin: "0 0 0.45rem", fontSize: "0.95rem" }}>Завести в 1С</h3>
          <ul style={{ margin: 0, paddingLeft: "1.15rem", display: "grid", gap: "0.35rem" }}>
            {toCreate.map((item) => {
              const kpLabel = formatKpIdsLabel(item.kp_ids);
              return (
                <li
                  key={`${item.product_kind}-${item.mark}-${item.field ?? "guid_1c"}`}
                  style={{ display: "grid", gap: "0.35rem" }}
                >
                  <div
                    style={{
                      display: "flex",
                      flexWrap: "wrap",
                      gap: "0.5rem",
                      alignItems: "center",
                    }}
                  >
                    <span>
                      {productKindLabel(item.product_kind)} · {item.mark} — {item.hint}
                      {kpLabel ? ` — ${kpLabel}` : ""}
                    </span>
                    {canOpenImport(item.product_kind) ? (
                      <Button
                        type="button"
                        aria-label={`Загрузить выгрузку 1С, ${item.mark}`}
                        onClick={onOpenImport}
                      >
                        Загрузить выгрузку 1С
                      </Button>
                    ) : null}
                  </div>
                  <GuidExplainPanel
                    source={economistGuidExplainSource}
                    productKind={item.product_kind}
                    mark={item.mark}
                    onCommitted={refreshTasks}
                  />
                </li>
              );
            })}
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
