import { useEffect, useState, type CSSProperties, type FormEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import { useGenerateGsmWaybillsMutation } from "@/features/gsm/hooks/useGsmQueries";
import { ApiError } from "@/shared/lib/apiError";
import type { FleetOverviewRow, WaybillGenerateResult } from "@/features/gsm/types/gsm";

type Props = {
  open: boolean;
  row: FleetOverviewRow | null;
  periodFrom: string;
  periodTo: string;
  onClose: () => void;
  onGenerated?: (result: WaybillGenerateResult) => void;
};

const field: CSSProperties = { display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467" };

const parseOptionalNumber = (raw: string): number | null => {
  const trimmed = raw.trim().replace(",", ".");
  if (!trimmed) return null;
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : Number.NaN;
};

export const VehicleGenerateDialog = ({
  open,
  row,
  periodFrom,
  periodTo,
  onClose,
  onGenerated,
}: Props) => {
  const generateMutation = useGenerateGsmWaybillsMutation();
  const [from, setFrom] = useState(periodFrom);
  const [to, setTo] = useState(periodTo);
  const [fuelStart, setFuelStart] = useState("");
  const [odometerStart, setOdometerStart] = useState("");
  const [force, setForce] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [startRequired, setStartRequired] = useState(false);
  const [result, setResult] = useState<WaybillGenerateResult | null>(null);

  useEffect(() => {
    if (open) {
      setFrom(periodFrom);
      setTo(periodTo);
      setFuelStart("");
      setOdometerStart("");
      setForce(false);
      setError(null);
      setStartRequired(false);
      setResult(null);
    }
  }, [open, periodFrom, periodTo, row?.vehicle.id]);

  const onSubmit = async (event: FormEvent) => {
    event.preventDefault();
    if (!row) return;
    setError(null);
    setStartRequired(false);
    const fuel = parseOptionalNumber(fuelStart);
    const odo = parseOptionalNumber(odometerStart);
    if (Number.isNaN(fuel) || (fuel != null && fuel < 0)) {
      setError("Стартовый остаток бака должен быть числом ≥ 0.");
      return;
    }
    if (Number.isNaN(odo) || (odo != null && (!Number.isInteger(odo) || odo < 0))) {
      setError("Стартовый одометр должен быть целым числом ≥ 0.");
      return;
    }
    try {
      const generated = await generateMutation.mutateAsync({
        vehicle_id: row.vehicle.id,
        period_from: from,
        period_to: to,
        force,
        fuel_start: fuel,
        odometer_start: odo == null ? null : Math.trunc(odo),
      });
      setResult(generated);
      onGenerated?.(generated);
    } catch (err) {
      if (err instanceof ApiError && err.code === "gsm_start_required") {
        setStartRequired(true);
      }
      setError(formatGsmError(err));
    }
  };

  const highlight = startRequired
    ? { outline: "2px solid #f04438", borderRadius: 12 }
    : undefined;

  return (
    <Modal open={open} onClose={onClose} title={row ? `Генерация: ${row.vehicle.name}` : "Генерация"}>
      <form onSubmit={onSubmit} style={{ display: "grid", gap: "0.85rem" }}>
        <label style={field}>
          С
          <Input type="date" aria-label="Период с" value={from} onChange={(e) => setFrom(e.target.value)} />
        </label>
        <label style={field}>
          По
          <Input type="date" aria-label="Период по" value={to} onChange={(e) => setTo(e.target.value)} />
        </label>
        <label style={field}>
          Старт бака, л
          <div style={highlight}>
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
        <label style={field}>
          Старт одометра
          <div style={highlight}>
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
        <label style={{ display: "flex", gap: 8, alignItems: "center", fontSize: "0.9rem" }}>
          <input
            type="checkbox"
            checked={force}
            onChange={(e) => setForce(e.target.checked)}
            aria-label="Перезаписать confirmed"
          />
          Перезаписать confirmed
        </label>
        {error && <Alert tone="error">{error}</Alert>}
        {result && (
          <Alert tone={result.manual_days > 0 ? "warning" : "success"}>
            Создано {result.days_created} дней, {result.manual_days} требуют ручной доработки.
            {result.problematic_days.length > 0 && (
              <ul aria-label="Дни ручной доработки">
                {result.problematic_days.map((day) => (
                  <li key={day.date}>
                    {day.date}: {day.detail}
                  </li>
                ))}
              </ul>
            )}
          </Alert>
        )}
        <div style={{ display: "flex", gap: "0.5rem", justifyContent: "flex-end" }}>
          <Button type="button" variant="secondary" onClick={onClose}>
            Закрыть
          </Button>
          <Button type="submit" disabled={generateMutation.isPending || !row}>
            {generateMutation.isPending ? "Генерация…" : "Сгенерировать"}
          </Button>
        </div>
      </form>
    </Modal>
  );
};
