import { useState, type CSSProperties } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import {
  useGsmCardsQuery,
  useGsmVehiclesQuery,
  usePatchCardMutation,
} from "@/features/gsm/hooks/useGsmQueries";
import type { GsmCard, GsmVehicle } from "@/features/gsm/types/gsm";

const sectionStyle: CSSProperties = {
  display: "grid",
  gap: "0.75rem",
  padding: "1rem",
  borderRadius: 12,
  border: "1px solid #e4e7ec",
  background: "#ffffff",
};

const thStyle: CSSProperties = { padding: "0.5rem", textAlign: "left", borderBottom: "1px solid #eaecf0" };
const tdStyle: CSSProperties = { padding: "0.5rem", borderBottom: "1px solid #f2f4f7", verticalAlign: "middle" };

const selectStyle: CSSProperties = {
  border: "1px solid #d0d5dd",
  borderRadius: 12,
  padding: "0.55rem 0.7rem",
  background: "#ffffff",
  minWidth: 180,
};

const vehicleLabel = (vehicles: GsmVehicle[], vehicleId: number | null): string => {
  if (vehicleId == null) {
    return "—";
  }
  const found = vehicles.find((v) => v.id === vehicleId);
  return found ? `${found.name} (${found.plate_number})` : `#${vehicleId}`;
};

export const CardsRegistryView = () => {
  const [includeArchived, setIncludeArchived] = useState(false);
  const cardsQuery = useGsmCardsQuery(includeArchived);
  const vehiclesQuery = useGsmVehiclesQuery(true);
  const patchMutation = usePatchCardMutation();
  const [actionError, setActionError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [pendingId, setPendingId] = useState<number | null>(null);

  const cards = cardsQuery.data ?? [];
  const vehicles = vehiclesQuery.data ?? [];
  const isLoading = cardsQuery.isLoading || vehiclesQuery.isLoading;
  const loadError = cardsQuery.error ?? vehiclesQuery.error;

  const bindVehicle = async (card: GsmCard, vehicleIdRaw: string) => {
    const vehicle_id = vehicleIdRaw === "" ? null : Number(vehicleIdRaw);
    if (vehicleIdRaw !== "" && !Number.isFinite(vehicle_id)) {
      setActionError("Выберите машину.");
      return;
    }
    setActionError(null);
    setInfo(null);
    setPendingId(card.id);
    try {
      await patchMutation.mutateAsync({
        id: card.id,
        payload: { vehicle_id },
      });
      setInfo(
        vehicle_id == null
          ? `Карта ${card.card_number}: привязка снята.`
          : `Карта ${card.card_number} привязана к машине.`,
      );
    } catch (err) {
      setActionError(formatGsmError(err));
    } finally {
      setPendingId(null);
    }
  };

  const archiveCard = async (card: GsmCard) => {
    setActionError(null);
    setInfo(null);
    setPendingId(card.id);
    try {
      await patchMutation.mutateAsync({
        id: card.id,
        payload: { archive: true },
      });
      setInfo(`Карта ${card.card_number} архивирована.`);
    } catch (err) {
      setActionError(formatGsmError(err));
    } finally {
      setPendingId(null);
    }
  };

  if (isLoading) {
    return (
      <div style={sectionStyle} aria-label="Топливные карты ГСМ">
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Топливные карты</h2>
        <Spinner />
      </div>
    );
  }

  if (loadError) {
    return (
      <div style={sectionStyle} aria-label="Топливные карты ГСМ">
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Топливные карты</h2>
        <Alert tone="error">{formatGsmError(loadError)}</Alert>
      </div>
    );
  }

  return (
    <div style={sectionStyle} aria-label="Топливные карты ГСМ">
      <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem", flexWrap: "wrap", alignItems: "center" }}>
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Топливные карты ({cards.length})</h2>
        <label style={{ display: "flex", gap: "0.45rem", alignItems: "center", color: "#475467" }}>
          <input
            type="checkbox"
            checked={includeArchived}
            onChange={(e) => setIncludeArchived(e.target.checked)}
          />
          Показывать архивные
        </label>
      </div>

      {info && <Alert tone="success">{info}</Alert>}
      {actionError && <Alert tone="error">{actionError}</Alert>}

      {cards.length === 0 ? (
        <Alert tone="info">Карты не найдены.</Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.92rem" }}>
            <thead>
              <tr>
                <th style={thStyle}>Номер</th>
                <th style={thStyle}>Машина</th>
                <th style={thStyle}>Назначена</th>
                <th style={thStyle}>Статус</th>
                <th style={thStyle}>Действия</th>
              </tr>
            </thead>
            <tbody>
              {cards.map((card) => {
                const archived = card.archived_at != null;
                const busy = pendingId === card.id && patchMutation.isPending;
                return (
                  <tr
                    key={card.id}
                    style={archived ? { opacity: 0.65, background: "#f9fafb" } : undefined}
                  >
                    <td style={tdStyle}>{card.card_number}</td>
                    <td style={tdStyle}>
                      {archived ? (
                        vehicleLabel(vehicles, card.vehicle_id)
                      ) : (
                        <select
                          aria-label={`Машина для карты ${card.card_number}`}
                          value={card.vehicle_id ?? ""}
                          disabled={busy}
                          style={selectStyle}
                          onChange={(e) => void bindVehicle(card, e.target.value)}
                        >
                          <option value="">Не привязана</option>
                          {vehicles.map((v) => (
                            <option key={v.id} value={v.id}>
                              {v.name} ({v.plate_number})
                            </option>
                          ))}
                        </select>
                      )}
                    </td>
                    <td style={tdStyle}>{card.assigned_at}</td>
                    <td style={tdStyle}>{archived ? `архив ${card.archived_at}` : "активна"}</td>
                    <td style={tdStyle}>
                      {!archived && (
                        <Button
                          type="button"
                          variant="danger"
                          disabled={busy}
                          onClick={() => void archiveCard(card)}
                        >
                          {busy ? "…" : "Архив"}
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
    </div>
  );
};
