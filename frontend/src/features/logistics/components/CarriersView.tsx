import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import { CarrierAutocomplete } from "@/features/logistics/components/CarrierAutocomplete";
import {
  useCarriersQuery,
  useMergeCarrierMutation,
} from "@/features/logistics/hooks/useLogisticsQueries";
import { isFlagOn } from "@/features/logistics/lib/logisticsFormat";
import { useDebouncedValue } from "@/features/logistics/lib/useDebouncedValue";
import type { Carrier } from "@/features/logistics/types/logistics";

export const CarriersView = () => {
  const [search, setSearch] = useState("");
  const debouncedSearch = useDebouncedValue(search.trim());
  const query = useCarriersQuery({ q: debouncedSearch });
  const carriers = query.data ?? [];

  const [mergeSource, setMergeSource] = useState<Carrier | null>(null);
  const [mergeTarget, setMergeTarget] = useState<{ id: number; name: string } | null>(null);
  const [mergeError, setMergeError] = useState<string | null>(null);
  const [mergeResult, setMergeResult] = useState<string | null>(null);
  const mergeMutation = useMergeCarrierMutation();

  const openMerge = (carrier: Carrier) => {
    setMergeSource(carrier);
    setMergeTarget(null);
    setMergeError(null);
  };

  const closeMerge = () => {
    setMergeSource(null);
    setMergeTarget(null);
    setMergeError(null);
  };

  const submitMerge = async () => {
    if (!mergeSource || !mergeTarget) {
      setMergeError("Выберите целевого перевозчика.");
      return;
    }
    if (mergeTarget.id === mergeSource.id) {
      setMergeError("Нельзя слить перевозчика самого с собой.");
      return;
    }
    setMergeError(null);
    try {
      const response = await mergeMutation.mutateAsync({
        id: mergeSource.id,
        intoId: mergeTarget.id,
      });
      setMergeResult(
        `«${mergeSource.name}» слит с «${mergeTarget.name}»: перенесено рейсов — ${response.moved_shipments}.`,
      );
      closeMerge();
    } catch (err) {
      setMergeError(getErrorMessage(err));
    }
  };

  return (
    <section style={{ display: "grid", gap: "1rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "0.75rem", flexWrap: "wrap" }}>
        <h2 style={{ margin: 0, fontSize: "1.25rem" }}>Перевозчики</h2>
        <div style={{ width: 320, maxWidth: "100%" }}>
          <Input
            type="text"
            placeholder="Поиск по названию"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
      </div>

      {mergeResult && <Alert tone="success">{mergeResult}</Alert>}
      {query.isError && <Alert tone="error">{getErrorMessage(query.error)}</Alert>}

      {query.isLoading ? (
        <Spinner />
      ) : carriers.length === 0 ? (
        <Alert tone="info">Перевозчики не найдены.</Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.92rem" }}>
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #eaecf0" }}>
                <th style={{ padding: "0.55rem" }}>Название</th>
                <th style={{ padding: "0.55rem" }}>Источник</th>
                <th style={{ padding: "0.55rem" }}>Рейсов</th>
                <th style={{ padding: "0.55rem" }}>Активен</th>
                <th style={{ padding: "0.55rem" }}>Действия</th>
              </tr>
            </thead>
            <tbody>
              {carriers.map((carrier) => {
                const active = isFlagOn(carrier.active);
                return (
                  <tr key={carrier.id} style={{ borderBottom: "1px solid #f2f4f7" }}>
                    <td style={{ padding: "0.55rem", fontWeight: 600 }}>{carrier.name}</td>
                    <td style={{ padding: "0.55rem" }}>{carrier.source_sheet || "—"}</td>
                    <td style={{ padding: "0.55rem" }}>{carrier.shipments_count}</td>
                    <td style={{ padding: "0.55rem" }}>{active ? "да" : "нет"}</td>
                    <td style={{ padding: "0.55rem" }}>
                      {active && (
                        <Button variant="secondary" onClick={() => openMerge(carrier)}>
                          Слить с…
                        </Button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      <Modal
        open={mergeSource !== null}
        onClose={closeMerge}
        title={mergeSource ? `Слить: ${mergeSource.name}` : "Слияние"}
        maxWidth={520}
      >
        {mergeSource && (
          <div style={{ display: "grid", gap: "0.75rem" }}>
            <p style={{ margin: 0, color: "#475467" }}>
              Все рейсы дубля ({mergeSource.shipments_count}) будут перенесены на выбранного
              перевозчика, дубль станет неактивным.
            </p>
            <label style={{ display: "grid", gap: "0.45rem" }}>
              <span style={{ fontWeight: 600 }}>Слить с перевозчиком</span>
              <CarrierAutocomplete
                selected={mergeTarget}
                onSelect={setMergeTarget}
                placeholder="Начните вводить название целевого"
              />
            </label>
            {mergeTarget && mergeTarget.id !== mergeSource.id && (
              <Alert tone="warning">
                Будет перенесено рейсов: {mergeSource.shipments_count} → «{mergeTarget.name}».
              </Alert>
            )}
            {mergeError && <Alert tone="error">{mergeError}</Alert>}
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button variant="ghost" onClick={closeMerge} disabled={mergeMutation.isPending}>
                Отмена
              </Button>
              <Button
                variant="danger"
                onClick={submitMerge}
                disabled={mergeMutation.isPending || !mergeTarget || mergeTarget.id === mergeSource.id}
              >
                {mergeMutation.isPending ? "Слияние..." : "Подтвердить слияние"}
              </Button>
            </div>
          </div>
        )}
      </Modal>
    </section>
  );
};
