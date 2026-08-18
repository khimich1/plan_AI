import { useEffect, useMemo, useState, type CSSProperties, type FormEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Drawer } from "@/shared/ui/Drawer";
import { Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import {
  formatKm,
  formatLiters,
  formatOdometer,
  formatRouteSummary,
} from "@/features/gsm/lib/waybillWarnings";
import {
  libraryRouteToLegs,
  previewDownstream,
} from "@/features/gsm/lib/downstreamPreview";
import {
  useGsmDriversQuery,
  useGsmRoutesQuery,
  useGsmSettingsQuery,
  useGsmStationsQuery,
  usePatchGsmWaybillMutation,
} from "@/features/gsm/hooks/useGsmQueries";
import type {
  GsmRoute,
  GsmStation,
  GsmVehicle,
  GsmWaybill,
  WaybillRouteLeg,
} from "@/features/gsm/types/gsm";

type Props = {
  open: boolean;
  waybill: GsmWaybill | null;
  vehicle: GsmVehicle | null;
  periodWaybills: GsmWaybill[];
  onClose: () => void;
  onSaved?: (waybill: GsmWaybill) => void;
};

const fieldGrid: CSSProperties = {
  display: "grid",
  gap: "0.85rem",
};

const labelStyle: CSSProperties = {
  display: "grid",
  gap: 4,
  fontSize: "0.85rem",
  color: "#475467",
};

const selectStyle: CSSProperties = {
  border: "1px solid #d0d5dd",
  borderRadius: 12,
  padding: "0.75rem 0.85rem",
  background: "#ffffff",
  width: "100%",
};

const previewTable: CSSProperties = {
  width: "100%",
  borderCollapse: "collapse",
  fontSize: "0.85rem",
};

const thTd: CSSProperties = {
  padding: "0.4rem 0.5rem",
  borderBottom: "1px solid #eaecf0",
  textAlign: "left",
};

const routeLabel = (route: GsmRoute, stationsById: Map<number, GsmStation>): string => {
  const stations = route.typical_station_ids
    .map((id) => stationsById.get(id)?.brand || stationsById.get(id)?.address)
    .filter(Boolean)
    .slice(0, 2);
  const azs = stations.length ? ` · АЗС: ${stations.join(", ")}` : "";
  return `${route.addr_a} → ${route.addr_b} (${route.km} км)${azs}`;
};

export const WaybillDayDrawer = ({
  open,
  waybill,
  vehicle,
  periodWaybills,
  onClose,
  onSaved,
}: Props) => {
  const vehicleId = waybill?.vehicle_id ?? null;
  const driversQuery = useGsmDriversQuery(true);
  const routesQuery = useGsmRoutesQuery(open ? vehicleId : null);
  const stationsQuery = useGsmStationsQuery();
  const settingsQuery = useGsmSettingsQuery();
  const patchMutation = usePatchGsmWaybillMutation();

  const [driverId, setDriverId] = useState<number | "">("");
  const [km, setKm] = useState("");
  const [routeId, setRouteId] = useState<number | "">("");
  const [stationFilter, setStationFilter] = useState<number | "">("");
  const [selectedRoute, setSelectedRoute] = useState<WaybillRouteLeg[] | null>(null);
  const [formError, setFormError] = useState<string | null>(null);

  useEffect(() => {
    if (!open || !waybill) {
      return;
    }
    setDriverId(waybill.driver_id);
    setKm(String(waybill.km));
    setRouteId(waybill.route.find((l) => l.route_id != null)?.route_id ?? "");
    setStationFilter("");
    setSelectedRoute(waybill.route.length ? waybill.route : null);
    setFormError(null);
    patchMutation.reset();
    // eslint-disable-next-line react-hooks/exhaustive-deps -- reset only when day opens
  }, [open, waybill?.id]);

  const stationsById = useMemo(() => {
    const map = new Map<number, GsmStation>();
    for (const s of stationsQuery.data ?? []) {
      map.set(s.id, s);
    }
    return map;
  }, [stationsQuery.data]);

  const routes = routesQuery.data ?? [];
  const filteredRoutes = useMemo(() => {
    if (stationFilter === "") {
      return routes;
    }
    return routes.filter((r) => r.typical_station_ids.includes(Number(stationFilter)));
  }, [routes, stationFilter]);

  const stationOptions = useMemo(() => {
    const ids = new Set<number>();
    for (const r of routes) {
      for (const id of r.typical_station_ids) {
        ids.add(id);
      }
    }
    return [...ids]
      .map((id) => stationsById.get(id))
      .filter((s): s is GsmStation => Boolean(s))
      .sort((a, b) => a.address.localeCompare(b.address, "ru"));
  }, [routes, stationsById]);

  const editedKm = Number(km);
  const preview = useMemo(() => {
    if (!waybill || !vehicle || !Number.isFinite(editedKm) || editedKm < 0) {
      return null;
    }
    return previewDownstream({
      edited: waybill,
      editedKm: Math.trunc(editedKm),
      periodWaybills,
      vehicle,
      winterStartMMDD: settingsQuery.data?.winter_start ?? "11-01",
    });
  }, [waybill, vehicle, editedKm, periodWaybills, settingsQuery.data?.winter_start]);

  const previewChanged = preview?.days.filter((d) => d.changed) ?? [];

  const applyLibraryRoute = (id: number) => {
    const route = routes.find((r) => r.id === id);
    if (!route) {
      return;
    }
    setRouteId(id);
    setKm(String(route.km));
    setSelectedRoute(libraryRouteToLegs(route));
  };

  const handleClose = () => {
    if (patchMutation.isPending) {
      return;
    }
    onClose();
  };

  const onSubmit = async (event: FormEvent) => {
    event.preventDefault();
    if (!waybill) {
      return;
    }
    setFormError(null);
    const kmNum = Number(km);
    if (!Number.isFinite(kmNum) || kmNum < 0 || !Number.isInteger(kmNum)) {
      setFormError("Км должны быть целым числом ≥ 0.");
      return;
    }
    if (driverId === "") {
      setFormError("Выберите водителя.");
      return;
    }
    if (preview?.error) {
      setFormError(preview.error);
      return;
    }

    const payload = {
      driver_id: Number(driverId),
      km: kmNum,
      route: selectedRoute,
    };

    try {
      const saved = await patchMutation.mutateAsync({ id: waybill.id, payload });
      onSaved?.(saved);
      onClose();
    } catch (err) {
      setFormError(formatGsmError(err));
    }
  };

  const title = waybill
    ? `Правка дня ${waybill.date}${vehicle ? ` · ${vehicle.name}` : ""}`
    : "Правка дня";

  return (
    <Drawer
      open={open && Boolean(waybill)}
      onClose={handleClose}
      title={title}
      width={560}
      footer={
        <div style={{ display: "flex", gap: "0.65rem", justifyContent: "flex-end" }}>
          <Button type="button" variant="secondary" onClick={handleClose} disabled={patchMutation.isPending}>
            Отмена
          </Button>
          <Button type="submit" form="waybill-day-form" disabled={patchMutation.isPending || Boolean(preview?.error)}>
            {patchMutation.isPending ? "Сохранение…" : "Сохранить"}
          </Button>
        </div>
      }
    >
      {!waybill ? null : (
        <form id="waybill-day-form" style={fieldGrid} onSubmit={onSubmit}>
          <div style={{ fontSize: "0.9rem", color: "#475467" }}>
            Текущий маршрут: {formatRouteSummary(waybill.route)} · {formatKm(waybill.km)}
          </div>

          <label style={labelStyle}>
            Фильтр по АЗС
            <select
              aria-label="Фильтр по АЗС"
              value={stationFilter === "" ? "" : String(stationFilter)}
              onChange={(e) => {
                const raw = e.target.value;
                setStationFilter(raw === "" ? "" : Number(raw));
              }}
              style={selectStyle}
            >
              <option value="">Все маршруты</option>
              {stationOptions.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.brand ? `${s.brand} — ${s.address}` : s.address}
                </option>
              ))}
            </select>
          </label>

          <label style={labelStyle}>
            Маршрут из библиотеки
            {routesQuery.isLoading ? (
              <Spinner />
            ) : (
              <select
                aria-label="Маршрут из библиотеки"
                value={routeId === "" ? "" : String(routeId)}
                onChange={(e) => {
                  const raw = e.target.value;
                  if (raw === "") {
                    setRouteId("");
                    return;
                  }
                  applyLibraryRoute(Number(raw));
                }}
                style={selectStyle}
              >
                <option value="">Не менять / текущий</option>
                {filteredRoutes.map((r) => (
                  <option key={r.id} value={r.id}>
                    {routeLabel(r, stationsById)}
                  </option>
                ))}
              </select>
            )}
            {routesQuery.error && (
              <Alert tone="error">{formatGsmError(routesQuery.error)}</Alert>
            )}
            {!routesQuery.isLoading && filteredRoutes.length === 0 && (
              <span style={{ color: "#667085" }}>Нет маршрутов для выбранного фильтра.</span>
            )}
          </label>

          <label style={labelStyle}>
            Водитель
            <select
              aria-label="Водитель"
              value={driverId === "" ? "" : String(driverId)}
              onChange={(e) => {
                const raw = e.target.value;
                setDriverId(raw === "" ? "" : Number(raw));
              }}
              style={selectStyle}
            >
              <option value="">Выберите…</option>
              {(driversQuery.data ?? []).map((d) => (
                <option key={d.id} value={d.id}>
                  {d.full_name}
                </option>
              ))}
            </select>
          </label>

          <label style={labelStyle}>
            Км за день
            <Input
              type="number"
              min={0}
              step={1}
              aria-label="Км за день"
              value={km}
              onChange={(e) => setKm(e.target.value)}
            />
          </label>

          {formError && <Alert tone="error">{formError}</Alert>}
          {preview?.error && <Alert tone="error">{preview.error}</Alert>}

          <section aria-label="Превью пересчёта" style={{ display: "grid", gap: "0.5rem" }}>
            <h3 style={{ margin: 0, fontSize: "1rem" }}>Превью остатка / одометра</h3>
            <p style={{ margin: 0, color: "#667085", fontSize: "0.85rem" }}>
              До сохранения: как изменятся топливо и одометр у этого и последующих draft-дней.
            </p>
            {previewChanged.length === 0 && !preview?.error && (
              <Alert tone="info">Изменений в цепочке нет — значения совпадают с текущими.</Alert>
            )}
            {previewChanged.length > 0 && (
              <div style={{ overflowX: "auto" }}>
                <table style={previewTable}>
                  <thead>
                    <tr>
                      <th style={thTd}>Дата</th>
                      <th style={thTd}>Км</th>
                      <th style={thTd}>Топливо</th>
                      <th style={thTd}>Одометр</th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewChanged.map((d) => (
                      <tr key={d.id} style={{ background: "#eff8ff" }}>
                        <td style={thTd}>{d.date}</td>
                        <td style={thTd}>{formatKm(d.km)}</td>
                        <td style={thTd}>
                          {formatLiters(d.fuel_start)} → {formatLiters(d.fuel_end)}
                        </td>
                        <td style={thTd}>
                          {formatOdometer(d.odometer_start)} → {formatOdometer(d.odometer_end)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        </form>
      )}
    </Drawer>
  );
};
