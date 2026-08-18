import { useMemo, useState, type CSSProperties, type FormEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { ManualWaybillDialog } from "@/features/gsm/components/ManualWaybillDialog";
import { VehiclePeriodStrip } from "@/features/gsm/components/VehiclePeriodStrip";
import { WaybillDayDrawer } from "@/features/gsm/components/WaybillDayDrawer";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import { exportConfirmMessages, exportDisabledReason } from "@/features/gsm/lib/exportGate";
import { warningMeta } from "@/features/gsm/lib/waybillWarnings";
import {
  useExportGsmWaybillsMutation,
  useGenerateGsmWaybillsMutation,
  useGsmDriversQuery,
  useGsmVehiclesQuery,
  useGsmWaybillsQuery,
} from "@/features/gsm/hooks/useGsmQueries";
import type {
  GsmDriver,
  GsmWaybill,
  WaybillGenerateResult,
  WaybillListParams,
  WaybillWarningCode,
} from "@/features/gsm/types/gsm";

const sectionStyle: CSSProperties = {
  display: "grid",
  gap: "1rem",
};

const controlsStyle: CSSProperties = {
  display: "flex",
  gap: "0.75rem",
  flexWrap: "wrap",
  alignItems: "flex-end",
  border: "1px solid #eaecf0",
  borderRadius: 14,
  background: "#ffffff",
  padding: "0.9rem 1rem",
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
  padding: "0.8rem 0.9rem",
  background: "#ffffff",
  minWidth: 220,
};

const parseOptionalNumber = (raw: string): number | null => {
  const trimmed = raw.trim().replace(",", ".");
  if (!trimmed) {
    return null;
  }
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : Number.NaN;
};

const formatGenerateSummary = (result: WaybillGenerateResult): string =>
  `Создано ${result.days_created} дней, ${result.manual_days} требуют ручной доработки`;

export const GsmPeriodView = () => {
  const vehiclesQuery = useGsmVehiclesQuery(true);
  const driversQuery = useGsmDriversQuery(true);
  const generateMutation = useGenerateGsmWaybillsMutation();
  const exportMutation = useExportGsmWaybillsMutation();

  const [vehicleId, setVehicleId] = useState<number | "">("");
  const [periodFrom, setPeriodFrom] = useState("");
  const [periodTo, setPeriodTo] = useState("");
  const [fuelStart, setFuelStart] = useState("");
  const [odometerStart, setOdometerStart] = useState("");
  const [force, setForce] = useState(false);
  const [formError, setFormError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [periodWarnings, setPeriodWarnings] = useState<WaybillWarningCode[]>([]);
  const [expandedWarning, setExpandedWarning] = useState<string | null>(null);
  const [problematicDates, setProblematicDates] = useState<string[]>([]);
  const [selectedDay, setSelectedDay] = useState<GsmWaybill | null>(null);
  const [manualOpen, setManualOpen] = useState(false);
  const [exportConfirm, setExportConfirm] = useState<string[] | null>(null);

  const listParams: WaybillListParams | null =
    vehicleId !== "" && periodFrom && periodTo
      ? { vehicleId: Number(vehicleId), periodFrom, periodTo }
      : null;

  const waybillsQuery = useGsmWaybillsQuery(listParams);

  const vehicles = vehiclesQuery.data ?? [];
  const drivers = driversQuery.data ?? [];
  const driversById = useMemo(() => {
    const map = new Map<number, GsmDriver>();
    for (const d of drivers) {
      map.set(d.id, d);
    }
    return map;
  }, [drivers]);

  const selectedVehicle = vehicles.find((v) => v.id === vehicleId) ?? null;
  const waybills = waybillsQuery.data ?? [];
  const exportBlockReason = !listParams
    ? "Выберите машину и период."
    : exportDisabledReason(waybills, periodWarnings);
  const exportDisabled = exportBlockReason != null || exportMutation.isPending;
  const showExportHint = Boolean(listParams && exportBlockReason);

  const onGenerate = async (event: FormEvent) => {
    event.preventDefault();
    setFormError(null);
    setInfo(null);
    setPeriodWarnings([]);
    setExpandedWarning(null);
    setProblematicDates([]);

    if (vehicleId === "" || !periodFrom || !periodTo) {
      setFormError("Выберите машину и период.");
      return;
    }
    if (periodTo < periodFrom) {
      setFormError("Дата «по» должна быть не раньше даты «с».");
      return;
    }

    const fuel = parseOptionalNumber(fuelStart);
    const odo = parseOptionalNumber(odometerStart);
    if (Number.isNaN(fuel) || (fuel != null && fuel < 0)) {
      setFormError("Стартовый остаток бака должен быть числом ≥ 0.");
      return;
    }
    if (Number.isNaN(odo) || (odo != null && (!Number.isInteger(odo) || odo < 0))) {
      setFormError("Стартовый одометр должен быть целым числом ≥ 0.");
      return;
    }

    try {
      const result = await generateMutation.mutateAsync({
        vehicle_id: Number(vehicleId),
        period_from: periodFrom,
        period_to: periodTo,
        force,
        fuel_start: fuel,
        odometer_start: odo == null ? null : Math.trunc(odo),
      });
      setPeriodWarnings(result.warnings);
      setProblematicDates(result.problematic_days.map((day) => day.date));
      setInfo(formatGenerateSummary(result));
      void waybillsQuery.refetch();
    } catch (err) {
      setFormError(formatGsmError(err));
    }
  };

  const runExport = async () => {
    if (vehicleId === "" || !periodFrom || !periodTo) {
      return;
    }
    setFormError(null);
    setInfo(null);
    try {
      await exportMutation.mutateAsync({
        vehicle_ids: [Number(vehicleId)],
        from: periodFrom,
        to: periodTo,
      });
      setInfo("Скачан zip с бланками путевых листов.");
      void waybillsQuery.refetch();
    } catch (err) {
      setFormError(formatGsmError(err));
    }
  };

  const onExportClick = () => {
    if (exportDisabled) {
      return;
    }
    const confirms = exportConfirmMessages(waybills, periodWarnings);
    if (confirms.length > 0) {
      setExportConfirm(confirms);
      return;
    }
    void runExport();
  };

  const loadingMeta = vehiclesQuery.isLoading || driversQuery.isLoading;
  const loadError = vehiclesQuery.error ?? driversQuery.error;

  if (loadingMeta) {
    return (
      <section style={sectionStyle} aria-label="Проверка периода ГСМ">
        <Spinner />
      </section>
    );
  }

  if (loadError) {
    return (
      <section style={sectionStyle} aria-label="Проверка периода ГСМ">
        <Alert tone="error">{formatGsmError(loadError)}</Alert>
      </section>
    );
  }

  return (
    <section style={sectionStyle} aria-label="Проверка периода ГСМ">
      <header>
        <h2 style={{ margin: 0, fontSize: "1.25rem" }}>Период × машина</h2>
        <p style={{ margin: "0.35rem 0 0", color: "#475467" }}>
          Сгенерируйте черновики ПЛ и проверьте маршрут, остаток бака и одометр по дням.
        </p>
      </header>

      <form style={controlsStyle} onSubmit={onGenerate}>
        <label style={labelStyle}>
          Машина
          <select
            aria-label="Машина"
            value={vehicleId === "" ? "" : String(vehicleId)}
            onChange={(e) => {
              const raw = e.target.value;
              setVehicleId(raw === "" ? "" : Number(raw));
            }}
            style={selectStyle}
          >
            <option value="">Выберите…</option>
            {vehicles.map((v) => (
              <option key={v.id} value={v.id}>
                {v.name} ({v.plate_number})
              </option>
            ))}
          </select>
        </label>
        <label style={labelStyle}>
          С
          <div style={{ width: 150 }}>
            <Input
              type="date"
              aria-label="Период с"
              value={periodFrom}
              onChange={(e) => setPeriodFrom(e.target.value)}
            />
          </div>
        </label>
        <label style={labelStyle}>
          По
          <div style={{ width: 150 }}>
            <Input
              type="date"
              aria-label="Период по"
              value={periodTo}
              onChange={(e) => setPeriodTo(e.target.value)}
            />
          </div>
        </label>
        <label style={labelStyle}>
          Старт бака, л
          <div style={{ width: 120 }}>
            <Input
              type="text"
              inputMode="decimal"
              placeholder="авто"
              aria-label="Стартовый остаток бака"
              value={fuelStart}
              onChange={(e) => setFuelStart(e.target.value)}
            />
          </div>
        </label>
        <label style={labelStyle}>
          Старт одометра
          <div style={{ width: 140 }}>
            <Input
              type="text"
              inputMode="numeric"
              placeholder="авто"
              aria-label="Стартовый одометр"
              value={odometerStart}
              onChange={(e) => setOdometerStart(e.target.value)}
            />
          </div>
        </label>
        <label
          style={{
            display: "flex",
            gap: "0.45rem",
            alignItems: "center",
            color: "#475467",
            fontSize: "0.9rem",
            paddingBottom: "0.55rem",
          }}
        >
          <input
            type="checkbox"
            checked={force}
            onChange={(e) => setForce(e.target.checked)}
            aria-label="Перезаписать confirmed"
          />
          Перезаписать confirmed
        </label>
        <Button type="submit" disabled={generateMutation.isPending}>
          {generateMutation.isPending ? "Генерация…" : "Сгенерировать"}
        </Button>
        <Button
          type="button"
          variant="secondary"
          disabled={vehicleId === ""}
          onClick={() => setManualOpen(true)}
        >
          Ручной ПЛ
        </Button>
        <span style={{ display: "grid", gap: 4 }}>
          <Button
            type="button"
            variant="secondary"
            disabled={exportDisabled}
            title={exportBlockReason ?? undefined}
            aria-describedby={showExportHint ? "gsm-export-hint" : undefined}
            onClick={onExportClick}
          >
            {exportMutation.isPending ? "Экспорт…" : "Экспорт zip"}
          </Button>
          {showExportHint && (
            <span id="gsm-export-hint" style={{ fontSize: "0.8rem", color: "#667085", maxWidth: 280 }}>
              {exportBlockReason}
            </span>
          )}
        </span>
      </form>

      {formError && <Alert tone="error">{formError}</Alert>}
      {info && (
        <Alert tone="success">
          <div>{info}</div>
          {problematicDates.length > 0 && (
            <ul
              aria-label="Дни ручной доработки"
              style={{ margin: "0.4rem 0 0", paddingLeft: "1.2rem" }}
            >
              {problematicDates.map((date) => (
                <li key={date}>{date}</li>
              ))}
            </ul>
          )}
        </Alert>
      )}

      {periodWarnings.length > 0 && (
        <div
          style={{ display: "flex", flexWrap: "wrap", gap: "0.5rem", alignItems: "flex-start" }}
          aria-label="Предупреждения периода"
        >
          {periodWarnings.map((code) => {
            const meta = warningMeta(code);
            const open = expandedWarning === code;
            const reason = meta.reason;
            return (
              <span key={code} style={{ display: "grid", gap: 4 }}>
                <button
                  type="button"
                  onClick={() => setExpandedWarning(open ? null : code)}
                  aria-expanded={open}
                  aria-label={`Предупреждение периода: ${meta.short}`}
                  title={reason}
                  style={{
                    border: "none",
                    borderRadius: 8,
                    padding: "0.35rem 0.6rem",
                    fontWeight: 600,
                    cursor: "pointer",
                    background: code === "unsolvable" ? "#fee2e2" : "#fef3c7",
                    color: code === "unsolvable" ? "#b42318" : "#92400e",
                  }}
                >
                  {meta.short}
                </button>
                {open && (
                  <span role="status" style={{ fontSize: "0.85rem", color: "#475467", maxWidth: 480 }}>
                    {reason}
                  </span>
                )}
              </span>
            );
          })}
        </div>
      )}

      {!listParams && (
        <Alert tone="info">Выберите машину и период, чтобы загрузить или сгенерировать ПЛ.</Alert>
      )}

      {listParams && waybillsQuery.isLoading && <Spinner />}
      {listParams && waybillsQuery.error && (
        <Alert tone="error">{formatGsmError(waybillsQuery.error)}</Alert>
      )}

      {listParams && selectedVehicle && !waybillsQuery.isLoading && !waybillsQuery.error && (
        <VehiclePeriodStrip
          vehicleName={selectedVehicle.name}
          plateNumber={selectedVehicle.plate_number}
          tankVolumeLiters={selectedVehicle.tank_volume_liters}
          waybills={waybills}
          driversById={driversById}
          onDayClick={setSelectedDay}
        />
      )}

      <WaybillDayDrawer
        open={selectedDay != null}
        waybill={selectedDay}
        vehicle={selectedVehicle}
        periodWaybills={waybills}
        onClose={() => setSelectedDay(null)}
        onSaved={() => {
          void waybillsQuery.refetch();
        }}
      />

      <ManualWaybillDialog
        open={manualOpen}
        onClose={() => setManualOpen(false)}
        defaultVehicleId={vehicleId === "" ? null : Number(vehicleId)}
        defaultDate={periodFrom}
        periodWaybills={waybills}
        onCreated={() => {
          void waybillsQuery.refetch();
        }}
      />

      <Modal
        open={exportConfirm != null}
        onClose={() => setExportConfirm(null)}
        title="Экспорт путевых листов"
        maxWidth={480}
      >
        <div style={{ display: "grid", gap: "0.85rem" }}>
          {exportConfirm?.map((message) => (
            <p key={message} style={{ margin: 0, color: "#475467" }}>
              {message}
            </p>
          ))}
          <div style={{ display: "flex", gap: "0.75rem", justifyContent: "flex-end" }}>
            <Button type="button" variant="secondary" onClick={() => setExportConfirm(null)}>
              Отмена
            </Button>
            <Button
              type="button"
              disabled={exportMutation.isPending}
              onClick={() => {
                setExportConfirm(null);
                void runExport();
              }}
            >
              Скачать
            </Button>
          </div>
        </div>
      </Modal>
    </section>
  );
};
