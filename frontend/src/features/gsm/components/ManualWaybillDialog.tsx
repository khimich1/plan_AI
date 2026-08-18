import { useEffect, useMemo, useState, type CSSProperties, type FormEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import {
  libraryRouteToLegs,
  previousWaybillChainStart,
} from "@/features/gsm/lib/downstreamPreview";
import {
  useCreateGsmWaybillMutation,
  useGsmDriversQuery,
  useGsmRoutesQuery,
  useGsmStationsQuery,
  useGsmVehiclesQuery,
} from "@/features/gsm/hooks/useGsmQueries";
import type {
  GsmRoute,
  GsmStation,
  GsmWaybill,
  WaybillCreatePayload,
} from "@/features/gsm/types/gsm";

type Props = {
  open: boolean;
  onClose: () => void;
  /** Prefill vehicle/period context from GsmPeriodView. */
  defaultVehicleId?: number | null;
  defaultDate?: string;
  periodWaybills?: GsmWaybill[];
  onCreated?: (waybill: GsmWaybill) => void;
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

const routeLabel = (route: GsmRoute, stationsById: Map<number, GsmStation>): string => {
  const stations = route.typical_station_ids
    .map((id) => stationsById.get(id)?.brand || stationsById.get(id)?.address)
    .filter(Boolean)
    .slice(0, 2);
  const azs = stations.length ? ` · АЗС: ${stations.join(", ")}` : "";
  return `${route.addr_a} → ${route.addr_b} (${route.km} км)${azs}`;
};

const parseOptionalNumber = (raw: string): number | null => {
  const trimmed = raw.trim().replace(",", ".");
  if (!trimmed) {
    return null;
  }
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : Number.NaN;
};

export const ManualWaybillDialog = ({
  open,
  onClose,
  defaultVehicleId = null,
  defaultDate = "",
  periodWaybills = [],
  onCreated,
}: Props) => {
  const vehiclesQuery = useGsmVehiclesQuery(true);
  const driversQuery = useGsmDriversQuery(true);
  const stationsQuery = useGsmStationsQuery();
  const createMutation = useCreateGsmWaybillMutation();

  const [vehicleId, setVehicleId] = useState<number | "">("");
  const [date, setDate] = useState("");
  const [driverId, setDriverId] = useState<number | "">("");
  const [routeId, setRouteId] = useState<number | "">("");
  const [stationFilter, setStationFilter] = useState<number | "">("");
  const [fuelStart, setFuelStart] = useState("");
  const [fuelIssued, setFuelIssued] = useState("0");
  const [odometerStart, setOdometerStart] = useState("");
  const [formError, setFormError] = useState<string | null>(null);
  const [autofillTouched, setAutofillTouched] = useState(false);

  const routesQuery = useGsmRoutesQuery(open && vehicleId !== "" ? Number(vehicleId) : null);

  useEffect(() => {
    if (!open) {
      return;
    }
    setVehicleId(defaultVehicleId ?? "");
    setDate(defaultDate);
    setDriverId("");
    setRouteId("");
    setStationFilter("");
    setFuelStart("");
    setFuelIssued("0");
    setOdometerStart("");
    setFormError(null);
    setAutofillTouched(false);
    createMutation.reset();
    // eslint-disable-next-line react-hooks/exhaustive-deps -- reset on open
  }, [open, defaultVehicleId, defaultDate]);

  useEffect(() => {
    if (!open || !date || autofillTouched) {
      return;
    }
    const chain = previousWaybillChainStart(periodWaybills, date);
    if (chain.fuel_start != null) {
      setFuelStart(String(chain.fuel_start));
    }
    if (chain.odometer_start != null) {
      setOdometerStart(String(chain.odometer_start));
    }
  }, [open, date, periodWaybills, autofillTouched]);

  useEffect(() => {
    if (!open || vehicleId === "" || driverId !== "") {
      return;
    }
    const vehicle = (vehiclesQuery.data ?? []).find((v) => v.id === vehicleId);
    if (vehicle?.primary_driver_id != null) {
      setDriverId(vehicle.primary_driver_id);
    }
  }, [open, vehicleId, vehiclesQuery.data, driverId]);

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

  const handleClose = () => {
    if (createMutation.isPending) {
      return;
    }
    onClose();
  };

  const onSubmit = async (event: FormEvent) => {
    event.preventDefault();
    setFormError(null);

    if (vehicleId === "" || !date || driverId === "" || routeId === "") {
      setFormError("Заполните машину, дату, водителя и маршрут.");
      return;
    }

    const route = routes.find((r) => r.id === Number(routeId));
    if (!route) {
      setFormError("Выберите маршрут из библиотеки.");
      return;
    }

    const fuel = parseOptionalNumber(fuelStart);
    const issued = parseOptionalNumber(fuelIssued);
    const odo = parseOptionalNumber(odometerStart);
    if (Number.isNaN(fuel) || (fuel != null && fuel < 0)) {
      setFormError("Остаток бака должен быть числом ≥ 0.");
      return;
    }
    if (Number.isNaN(issued) || issued == null || issued < 0) {
      setFormError("Выдано, л — число ≥ 0.");
      return;
    }
    if (Number.isNaN(odo) || (odo != null && (!Number.isInteger(odo) || odo < 0))) {
      setFormError("Одометр должен быть целым числом ≥ 0.");
      return;
    }

    const payload: WaybillCreatePayload = {
      vehicle_id: Number(vehicleId),
      date,
      driver_id: Number(driverId),
      route: libraryRouteToLegs(route),
      fuel_issued: issued,
      fuel_start: fuel,
      odometer_start: odo == null ? null : Math.trunc(odo),
    };

    try {
      const created = await createMutation.mutateAsync(payload);
      onCreated?.(created);
      onClose();
    } catch (err) {
      setFormError(formatGsmError(err));
    }
  };

  return (
    <Modal open={open} onClose={handleClose} title="Ручной путевой лист" maxWidth={560}>
      <form style={fieldGrid} onSubmit={onSubmit}>
        <label style={labelStyle}>
          Машина
          <select
            aria-label="Машина"
            value={vehicleId === "" ? "" : String(vehicleId)}
            onChange={(e) => {
              const raw = e.target.value;
              setVehicleId(raw === "" ? "" : Number(raw));
              setRouteId("");
              setStationFilter("");
            }}
            style={selectStyle}
          >
            <option value="">Выберите…</option>
            {(vehiclesQuery.data ?? []).map((v) => (
              <option key={v.id} value={v.id}>
                {v.name} ({v.plate_number})
              </option>
            ))}
          </select>
        </label>

        <label style={labelStyle}>
          Дата
          <Input
            type="date"
            aria-label="Дата"
            value={date}
            onChange={(e) => {
              setDate(e.target.value);
              setAutofillTouched(false);
            }}
          />
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
          Фильтр по АЗС
          <select
            aria-label="Фильтр по АЗС"
            value={stationFilter === "" ? "" : String(stationFilter)}
            onChange={(e) => {
              const raw = e.target.value;
              setStationFilter(raw === "" ? "" : Number(raw));
            }}
            style={selectStyle}
            disabled={vehicleId === ""}
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
          Маршрут
          {routesQuery.isLoading ? (
            <Spinner />
          ) : (
            <select
              aria-label="Маршрут"
              value={routeId === "" ? "" : String(routeId)}
              onChange={(e) => {
                const raw = e.target.value;
                setRouteId(raw === "" ? "" : Number(raw));
              }}
              style={selectStyle}
              disabled={vehicleId === ""}
            >
              <option value="">Выберите…</option>
              {filteredRoutes.map((r) => (
                <option key={r.id} value={r.id}>
                  {routeLabel(r, stationsById)}
                </option>
              ))}
            </select>
          )}
        </label>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
            gap: "0.75rem",
          }}
        >
          <label style={labelStyle}>
            Остаток бака, л
            <Input
              type="text"
              inputMode="decimal"
              aria-label="Остаток бака"
              value={fuelStart}
              onChange={(e) => {
                setAutofillTouched(true);
                setFuelStart(e.target.value);
              }}
              placeholder="из предыдущего дня"
            />
          </label>
          <label style={labelStyle}>
            Выдано, л
            <Input
              type="text"
              inputMode="decimal"
              aria-label="Выдано"
              value={fuelIssued}
              onChange={(e) => setFuelIssued(e.target.value)}
            />
          </label>
          <label style={labelStyle}>
            Одометр старт
            <Input
              type="text"
              inputMode="numeric"
              aria-label="Одометр старт"
              value={odometerStart}
              onChange={(e) => {
                setAutofillTouched(true);
                setOdometerStart(e.target.value);
              }}
              placeholder="из предыдущего дня"
            />
          </label>
        </div>

        <p style={{ margin: 0, fontSize: "0.85rem", color: "#667085" }}>
          Поля топлива и одометра подставляются из предыдущего ПЛ в периоде; их можно править вручную.
        </p>

        {formError && <Alert tone="error">{formError}</Alert>}

        <div style={{ display: "flex", gap: "0.65rem", justifyContent: "flex-end" }}>
          <Button type="button" variant="secondary" onClick={handleClose} disabled={createMutation.isPending}>
            Отмена
          </Button>
          <Button type="submit" disabled={createMutation.isPending}>
            {createMutation.isPending ? "Создание…" : "Создать ПЛ"}
          </Button>
        </div>
      </form>
    </Modal>
  );
};
