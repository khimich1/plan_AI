import { useState, type CSSProperties, type FormEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import {
  useGsmVehiclesQuery,
  usePatchVehicleMutation,
} from "@/features/gsm/hooks/useGsmQueries";
import type { GsmVehicle } from "@/features/gsm/types/gsm";

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

type Draft = {
  tank_volume_liters: string;
  norm_summer: string;
  norm_winter: string;
};

const toDraft = (v: GsmVehicle): Draft => ({
  tank_volume_liters: String(v.tank_volume_liters),
  norm_summer: String(v.norm_summer),
  norm_winter: String(v.norm_winter),
});

const parsePositive = (raw: string, label: string): number | string => {
  const normalized = raw.trim().replace(",", ".");
  const n = Number(normalized);
  if (!Number.isFinite(n) || n <= 0) {
    return `${label} должна быть числом больше 0.`;
  }
  return n;
};

export const VehiclesCard = () => {
  const vehiclesQuery = useGsmVehiclesQuery(true);
  const patchMutation = usePatchVehicleMutation();
  const [drafts, setDrafts] = useState<Record<number, Draft>>({});
  const [editingId, setEditingId] = useState<number | null>(null);
  const [rowError, setRowError] = useState<string | null>(null);
  const [savedOk, setSavedOk] = useState(false);

  const vehicles = vehiclesQuery.data ?? [];

  const startEdit = (vehicle: GsmVehicle) => {
    setEditingId(vehicle.id);
    setDrafts((prev) => ({ ...prev, [vehicle.id]: toDraft(vehicle) }));
    setRowError(null);
    setSavedOk(false);
  };

  const cancelEdit = () => {
    setEditingId(null);
    setRowError(null);
  };

  const updateDraft = (id: number, field: keyof Draft, value: string) => {
    setDrafts((prev) => ({
      ...prev,
      [id]: {
        ...(prev[id] ?? { tank_volume_liters: "", norm_summer: "", norm_winter: "" }),
        [field]: value,
      },
    }));
  };

  const saveVehicle = async (event: FormEvent, vehicle: GsmVehicle) => {
    event.preventDefault();
    const draft = drafts[vehicle.id] ?? toDraft(vehicle);
    const tank = parsePositive(draft.tank_volume_liters, "Ёмкость бака");
    if (typeof tank === "string") {
      setRowError(tank);
      return;
    }
    const summer = parsePositive(draft.norm_summer, "Летняя норма");
    if (typeof summer === "string") {
      setRowError(summer);
      return;
    }
    const winter = parsePositive(draft.norm_winter, "Зимняя норма");
    if (typeof winter === "string") {
      setRowError(winter);
      return;
    }
    setRowError(null);
    try {
      await patchMutation.mutateAsync({
        id: vehicle.id,
        payload: {
          tank_volume_liters: tank,
          norm_summer: summer,
          norm_winter: winter,
        },
      });
      setEditingId(null);
      setSavedOk(true);
    } catch (err) {
      setRowError(formatGsmError(err));
    }
  };

  if (vehiclesQuery.isLoading) {
    return (
      <div style={sectionStyle} aria-label="Машины ГСМ">
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Машины</h2>
        <Spinner />
      </div>
    );
  }

  if (vehiclesQuery.error) {
    return (
      <div style={sectionStyle} aria-label="Машины ГСМ">
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Машины</h2>
        <Alert tone="error">{formatGsmError(vehiclesQuery.error)}</Alert>
      </div>
    );
  }

  return (
    <div style={sectionStyle} aria-label="Машины ГСМ">
      <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Машины ({vehicles.length})</h2>
      {savedOk && editingId == null && <Alert tone="success">Нормы и бак сохранены.</Alert>}
      {vehicles.length === 0 ? (
        <Alert tone="info">Машины не найдены.</Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.92rem" }}>
            <thead>
              <tr>
                <th style={thStyle}>Название</th>
                <th style={thStyle}>Госномер</th>
                <th style={thStyle}>Бак, л</th>
                <th style={thStyle}>Норма лето</th>
                <th style={thStyle}>Норма зима</th>
                <th style={thStyle}>Действия</th>
              </tr>
            </thead>
            <tbody>
              {vehicles.map((vehicle) => {
                const isEditing = editingId === vehicle.id;
                const draft = drafts[vehicle.id] ?? toDraft(vehicle);
                return (
                  <tr key={vehicle.id}>
                    <td style={tdStyle}>{vehicle.name}</td>
                    <td style={tdStyle}>{vehicle.plate_number}</td>
                    {isEditing ? (
                      <>
                        <td style={tdStyle}>
                          <Input
                            type="number"
                            min={0.1}
                            step="0.1"
                            value={draft.tank_volume_liters}
                            aria-label={`Бак ${vehicle.name}`}
                            onChange={(e) => updateDraft(vehicle.id, "tank_volume_liters", e.target.value)}
                          />
                        </td>
                        <td style={tdStyle}>
                          <Input
                            type="number"
                            min={0.1}
                            step="0.1"
                            value={draft.norm_summer}
                            aria-label={`Норма лето ${vehicle.name}`}
                            onChange={(e) => updateDraft(vehicle.id, "norm_summer", e.target.value)}
                          />
                        </td>
                        <td style={tdStyle}>
                          <Input
                            type="number"
                            min={0.1}
                            step="0.1"
                            value={draft.norm_winter}
                            aria-label={`Норма зима ${vehicle.name}`}
                            onChange={(e) => updateDraft(vehicle.id, "norm_winter", e.target.value)}
                          />
                        </td>
                        <td style={tdStyle}>
                          <form
                            onSubmit={(e) => void saveVehicle(e, vehicle)}
                            style={{ display: "flex", gap: "0.4rem", flexWrap: "wrap" }}
                          >
                            <Button type="submit" disabled={patchMutation.isPending}>
                              {patchMutation.isPending && editingId === vehicle.id
                                ? "Сохранение…"
                                : "Сохранить"}
                            </Button>
                            <Button
                              type="button"
                              variant="ghost"
                              onClick={cancelEdit}
                              disabled={patchMutation.isPending}
                            >
                              Отмена
                            </Button>
                          </form>
                        </td>
                      </>
                    ) : (
                      <>
                        <td style={tdStyle}>{vehicle.tank_volume_liters}</td>
                        <td style={tdStyle}>{vehicle.norm_summer}</td>
                        <td style={tdStyle}>{vehicle.norm_winter}</td>
                        <td style={tdStyle}>
                          <Button type="button" variant="secondary" onClick={() => startEdit(vehicle)}>
                            Нормы / бак
                          </Button>
                        </td>
                      </>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {rowError && <Alert tone="error">{rowError}</Alert>}
    </div>
  );
};
