import { useEffect, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { getErrorMessage } from "@/shared/lib/apiError";
import { logisticsApi } from "@/features/logistics/api/logisticsApi";
import {
  useCreateShipmentMutation,
  useReuseTransportMutation,
} from "@/features/logistics/hooks/useLogisticsQueries";
import {
  deliveryTypeLabel,
  todayIsoDate,
} from "@/features/logistics/lib/logisticsFormat";
import type { DeliveryType, LogisticsKpSearchItem } from "@/features/logistics/types/logistics";

type PickedKp = {
  kp_id: number;
  customer_name: string | null;
};

type Props = {
  open: boolean;
  onClose: () => void;
  onCreated: (shipmentId: number) => void;
  /** Если задан — submit идёт в reuse-transport вместо create. */
  sourceShipmentId?: number;
  /** Префилл типа Д/С при открытии (из строки-источника). */
  initialDeliveryType?: DeliveryType;
};

export const CreateShipmentDialog = ({
  open,
  onClose,
  onCreated,
  sourceShipmentId,
  initialDeliveryType,
}: Props) => {
  const [date, setDate] = useState(todayIsoDate());
  const [deliveryType, setDeliveryType] = useState<DeliveryType>("delivery");
  const [picked, setPicked] = useState<PickedKp[]>([]);
  const [searchText, setSearchText] = useState("");
  const [searchResults, setSearchResults] = useState<LogisticsKpSearchItem[]>([]);
  const [searchError, setSearchError] = useState<string | null>(null);
  const [searching, setSearching] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);

  const createMutation = useCreateShipmentMutation();
  const reuseMutation = useReuseTransportMutation();
  const isReuse = sourceShipmentId != null;
  const isPending = createMutation.isPending || reuseMutation.isPending;

  const resetForm = (nextDeliveryType: DeliveryType = "delivery") => {
    setDate(todayIsoDate());
    setDeliveryType(nextDeliveryType);
    setPicked([]);
    setSearchText("");
    setSearchResults([]);
    setSearchError(null);
    setSubmitError(null);
  };

  useEffect(() => {
    if (!open) {
      return;
    }
    resetForm(initialDeliveryType ?? "delivery");
  }, [open, initialDeliveryType, sourceShipmentId]);

  const close = () => {
    resetForm();
    onClose();
  };

  const addKp = (kp: PickedKp) => {
    setPicked((prev) => (prev.some((p) => p.kp_id === kp.kp_id) ? prev : [...prev, kp]));
  };

  const removeKp = (kpId: number) => {
    setPicked((prev) => prev.filter((p) => p.kp_id !== kpId));
  };

  const runSearch = async (event: React.FormEvent) => {
    event.preventDefault();
    const raw = searchText.trim();
    setSearchError(null);
    setSearchResults([]);
    if (!raw) {
      return;
    }
    setSearching(true);
    try {
      const numeric = Number(raw);
      const response = await logisticsApi.searchKp(
        Number.isInteger(numeric) && numeric > 0
          ? { kpId: numeric }
          : { customer: raw },
      );
      setSearchResults(response.items);
      if (response.items.length === 0) {
        setSearchError(
          "КП не найдено. Для поиска КП должно быть в статусе «в работе» или «На СГП».",
        );
      }
    } catch (error) {
      setSearchError(getErrorMessage(error));
    } finally {
      setSearching(false);
    }
  };

  const addByNumber = () => {
    const raw = searchText.trim();
    const kpId = Number(raw);
    if (!Number.isInteger(kpId) || kpId <= 0) {
      setSearchError("Укажите номер КП (целое число больше 0), чтобы добавить напрямую.");
      return;
    }
    setSearchError(null);
    addKp({ kp_id: kpId, customer_name: null });
    setSearchText("");
    setSearchResults([]);
  };

  const submit = async () => {
    setSubmitError(null);
    if (!date) {
      setSubmitError("Укажите дату рейса.");
      return;
    }
    if (picked.length === 0) {
      setSubmitError("Добавьте хотя бы один заказ (КП).");
      return;
    }
    const payload = {
      shipment_date: date,
      delivery_type: deliveryType,
      kp_ids: picked.map((p) => p.kp_id),
    };
    try {
      const created =
        isReuse && sourceShipmentId != null
          ? await reuseMutation.mutateAsync({ sourceId: sourceShipmentId, payload })
          : await createMutation.mutateAsync(payload);
      resetForm();
      onCreated(created.id);
    } catch (error) {
      setSubmitError(getErrorMessage(error));
    }
  };

  const title = isReuse ? "Новый рейс на основе" : "Новый рейс";

  return (
    <Modal open={open} onClose={close} title={title} maxWidth={640}>
      <div style={{ display: "grid", gap: "1rem" }}>
        {isReuse && (
          <Alert tone="info">
            Транспорт (перевозчик, водитель, ТС, доверенность) будет скопирован из рейса №
            {sourceShipmentId}. УПД, № заявки и состав не копируются.
          </Alert>
        )}
        <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
          <label style={{ display: "grid", gap: "0.45rem", flex: "1 1 180px" }}>
            <span style={{ fontWeight: 600 }}>Дата рейса</span>
            <Input type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          </label>
          <label style={{ display: "grid", gap: "0.45rem", flex: "1 1 180px" }}>
            <span style={{ fontWeight: 600 }}>Тип выдачи</span>
            <select
              value={deliveryType}
              onChange={(e) => setDeliveryType(e.target.value as DeliveryType)}
              style={{
                border: "1px solid #d0d5dd",
                borderRadius: 12,
                padding: "0.8rem 0.9rem",
                background: "#ffffff",
              }}
            >
              <option value="delivery">{deliveryTypeLabel("delivery")}</option>
              <option value="pickup">{deliveryTypeLabel("pickup")}</option>
            </select>
          </label>
        </div>

        <div style={{ display: "grid", gap: "0.5rem" }}>
          <span style={{ fontWeight: 600 }}>Заказы (КП)</span>
          <form onSubmit={runSearch} style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            <div style={{ flex: "1 1 220px" }}>
              <Input
                type="text"
                placeholder="Номер КП или заказчик"
                value={searchText}
                onChange={(e) => setSearchText(e.target.value)}
              />
            </div>
            <Button type="submit" variant="secondary" disabled={searching}>
              {searching ? "Поиск..." : "Найти"}
            </Button>
            <Button type="button" variant="ghost" onClick={addByNumber}>
              Добавить по номеру
            </Button>
          </form>
          {searchError && <Alert tone="warning">{searchError}</Alert>}
          {searchResults.length > 0 && (
            <ul
              style={{
                margin: 0,
                padding: 0,
                listStyle: "none",
                display: "grid",
                gap: "0.35rem",
                maxHeight: 180,
                overflow: "auto",
              }}
            >
              {searchResults.map((item) => {
                const already = picked.some((p) => p.kp_id === item.kp_id);
                return (
                  <li
                    key={item.kp_id}
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      gap: "0.5rem",
                      border: "1px solid #eaecf0",
                      borderRadius: 10,
                      padding: "0.45rem 0.65rem",
                    }}
                  >
                    <span>
                      КП №{item.kp_id}
                      {item.customer_name ? ` — ${item.customer_name}` : ""}
                    </span>
                    <Button
                      type="button"
                      variant="secondary"
                      disabled={already}
                      onClick={() =>
                        addKp({ kp_id: item.kp_id, customer_name: item.customer_name })
                      }
                    >
                      {already ? "Добавлен" : "Добавить"}
                    </Button>
                  </li>
                );
              })}
            </ul>
          )}
          {picked.length > 0 && (
            <div style={{ display: "flex", gap: "0.4rem", flexWrap: "wrap" }}>
              {picked.map((kp) => (
                <span
                  key={kp.kp_id}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: "0.4rem",
                    background: "#eef4ff",
                    color: "#1d4ed8",
                    borderRadius: 999,
                    padding: "0.25rem 0.7rem",
                    fontWeight: 600,
                    fontSize: "0.9rem",
                  }}
                >
                  КП №{kp.kp_id}
                  {kp.customer_name ? ` · ${kp.customer_name}` : ""}
                  <button
                    type="button"
                    aria-label={`Убрать КП №${kp.kp_id}`}
                    onClick={() => removeKp(kp.kp_id)}
                    style={{
                      border: "none",
                      background: "transparent",
                      cursor: "pointer",
                      color: "inherit",
                      fontWeight: 700,
                    }}
                  >
                    ×
                  </button>
                </span>
              ))}
            </div>
          )}
        </div>

        {submitError && <Alert tone="error">{submitError}</Alert>}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" onClick={close} disabled={isPending}>
            Отмена
          </Button>
          <Button onClick={submit} disabled={isPending}>
            {isPending ? "Создание..." : isReuse ? "Создать на основе" : "Создать рейс"}
          </Button>
        </div>
      </div>
    </Modal>
  );
};
