import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import {
  useGuidDuplicatesQuery,
  useResolveDuplicateMutation,
} from "@/features/nomenclature/hooks/useNomenclatureQueries";
import { productKindLabel, type DuplicateTask } from "@/features/nomenclature/types/nomenclature";
import { ApiError, getErrorMessage } from "@/shared/lib/apiError";

const formatPrice = (value: number | null): string => {
  if (value == null) {
    return "—";
  }
  return value.toLocaleString("ru-RU", { maximumFractionDigits: 2 });
};

const DuplicateCard = ({ item }: { item: DuplicateTask }) => {
  const resolveDuplicate = useResolveDuplicateMutation();
  const [chosen, setChosen] = useState(item.candidates[0]?.guid ?? "");
  const [note, setNote] = useState("");
  const [error, setError] = useState<string | null>(null);

  const onRemember = async () => {
    if (!chosen) {
      setError("Выберите GUID.");
      return;
    }
    setError(null);
    try {
      await resolveDuplicate.mutateAsync({
        scope: item.scope,
        key: item.key,
        chosen_guid: chosen,
        note,
        product_kind: item.product_kind,
      });
    } catch (err) {
      if (err instanceof ApiError && err.status === 409) {
        setError("кандидат исчез, задача открыта");
        return;
      }
      setError(getErrorMessage(err));
    }
  };

  return (
    <div
      style={{
        border: "1px solid #e4e7ec",
        borderRadius: 12,
        padding: "0.85rem 1rem",
        display: "grid",
        gap: "0.65rem",
      }}
    >
      <div style={{ fontWeight: 600 }}>
        {productKindLabel(item.product_kind)} · {item.key}
      </div>
      <fieldset style={{ border: "none", margin: 0, padding: 0, display: "grid", gap: "0.35rem" }}>
        <legend style={{ fontSize: "0.9rem", color: "#475467", padding: 0 }}>Кандидаты</legend>
        {item.candidates.map((candidate) => (
          <label key={candidate.guid} style={{ display: "flex", gap: "0.5rem", alignItems: "flex-start" }}>
            <input
              type="radio"
              name={`dup-${item.scope}-${item.key}`}
              value={candidate.guid}
              checked={chosen === candidate.guid}
              onChange={() => setChosen(candidate.guid)}
            />
            <span>
              {candidate.name} · {candidate.guid}
              <span style={{ color: "#667085" }}> · цена 1С {formatPrice(candidate.price)}</span>
            </span>
          </label>
        ))}
      </fieldset>
      <label style={{ display: "grid", gap: "0.35rem" }}>
        <span style={{ fontSize: "0.9rem", color: "#475467" }}>Заметка</span>
        <input
          value={note}
          onChange={(event) => setNote(event.target.value)}
          placeholder="по решению бухгалтерии"
          aria-label={`Заметка для ${item.key}`}
          style={{
            border: "1px solid #d0d5dd",
            borderRadius: 8,
            padding: "0.45rem 0.6rem",
          }}
        />
      </label>
      {error ? <Alert tone="error">{error}</Alert> : null}
      <div>
        <Button type="button" onClick={() => void onRemember()} disabled={resolveDuplicate.isPending}>
          Запомнить выбор
        </Button>
      </div>
    </div>
  );
};

export const DuplicatesSection = () => {
  const query = useGuidDuplicatesQuery();
  const items = query.data?.duplicates ?? [];

  return (
    <section aria-label="Дубли GUID" style={{ display: "grid", gap: "0.85rem" }}>
      <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Дубли GUID</h2>
      {query.isError ? <Alert tone="error">{getErrorMessage(query.error)}</Alert> : null}
      {items.length === 0 && !query.isLoading ? (
        <p style={{ margin: 0, color: "#667085" }}>дублей нет</p>
      ) : null}
      {items.map((item) => (
        <DuplicateCard key={`${item.scope}-${item.product_kind}-${item.key}`} item={item} />
      ))}
    </section>
  );
};
