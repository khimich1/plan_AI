import { useMemo, useRef, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import { ManualWaybillDialog } from "@/features/gsm/components/ManualWaybillDialog";
import { VehicleDayFeed } from "@/features/gsm/components/VehicleDayFeed";
import { VehicleMonthCalendar } from "@/features/gsm/components/VehicleMonthCalendar";
import { VehiclePeriodStrip } from "@/features/gsm/components/VehiclePeriodStrip";
import { WaybillDayDrawer } from "@/features/gsm/components/WaybillDayDrawer";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import { buildVehicleDayCells } from "@/features/gsm/lib/vehicleDayMap";
import {
  buildVehicleDayFeed,
  monthBounds,
  shiftMonth,
} from "@/features/gsm/lib/vehicleDayFeed";
import {
  useGsmDriversQuery,
  useGsmTransactionsQuery,
  useGsmVehiclesQuery,
  useGsmWaybillsQuery,
} from "@/features/gsm/hooks/useGsmQueries";
import type { GsmDriver, GsmWaybill } from "@/features/gsm/types/gsm";

type PeriodBounds = { from: string; to: string };

type Props = {
  vehicleId: number;
  vehicleName: string;
  plateNumber: string;
  periodFrom: string;
  periodTo: string;
  onGenerate?: (bounds: PeriodBounds) => void;
  onPeriodChange?: (bounds: PeriodBounds) => void;
};

export const VehicleWaybillJournal = ({
  vehicleId,
  vehicleName,
  plateNumber,
  periodFrom,
  periodTo,
  onGenerate,
  onPeriodChange,
}: Props) => {
  const displayMonth = periodFrom.slice(0, 7);
  const waybillsQuery = useGsmWaybillsQuery({
    vehicleId,
    periodFrom,
    periodTo,
  });
  const txQuery = useGsmTransactionsQuery({
    vehicleId,
    periodFrom,
    periodTo,
  });
  const driversQuery = useGsmDriversQuery(true);
  const vehiclesQuery = useGsmVehiclesQuery(true);
  const [selected, setSelected] = useState<GsmWaybill | null>(null);
  const [manualOpen, setManualOpen] = useState(false);
  const [rechainedNote, setRechainedNote] = useState<string | null>(null);
  const generateRef = useRef<HTMLButtonElement>(null);

  const waybills = waybillsQuery.data ?? [];
  const vehicle = (vehiclesQuery.data ?? []).find((item) => item.id === vehicleId) ?? null;
  const driversById = useMemo(() => {
    const map = new Map<number, GsmDriver>();
    for (const driver of driversQuery.data ?? []) {
      map.set(driver.id, driver);
    }
    return map;
  }, [driversQuery.data]);

  const txs = txQuery.error ? [] : (txQuery.data?.rows ?? []);

  const dayCells = useMemo(
    () => buildVehicleDayCells(periodFrom, periodTo, waybills, txs),
    [periodFrom, periodTo, waybills, txs],
  );

  const feed = useMemo(
    () => buildVehicleDayFeed(periodFrom, periodTo, waybills, txs),
    [periodFrom, periodTo, waybills, txs],
  );

  const focusGenerate = () => {
    const btn = generateRef.current;
    if (!btn) return;
    btn.focus();
    btn.scrollIntoView({ behavior: "smooth", block: "nearest" });
  };

  if (waybillsQuery.isLoading) {
    return <Spinner />;
  }
  if (waybillsQuery.error) {
    return <Alert tone="error">{formatGsmError(waybillsQuery.error)}</Alert>;
  }

  return (
    <div style={{ display: "grid", gap: "0.75rem" }} data-testid={`journal-${vehicleId}`}>
      <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem", flexWrap: "wrap" }}>
        <strong>Журнал ПЛ</strong>
        <span style={{ display: "flex", gap: "0.5rem" }}>
          {onGenerate && (
            <Button
              ref={generateRef}
              type="button"
              data-testid="journal-generate-btn"
              onClick={() => onGenerate({ from: periodFrom, to: periodTo })}
            >
              Сгенерировать
            </Button>
          )}
          <Button type="button" variant="secondary" onClick={() => setManualOpen(true)}>
            + Ручной ПЛ
          </Button>
        </span>
      </div>

      {txQuery.error && (
        <Alert tone="error">{formatGsmError(txQuery.error)}</Alert>
      )}

      {rechainedNote && <Alert tone="success">{rechainedNote}</Alert>}

      {txQuery.isLoading ? (
        <Spinner />
      ) : (
        <VehicleMonthCalendar
          cells={dayCells}
          month={displayMonth}
          onMonthChange={(delta) =>
            onPeriodChange?.(monthBounds(shiftMonth(displayMonth, delta)))
          }
          onDayClick={setSelected}
          onGapClick={focusGenerate}
        />
      )}

      <VehiclePeriodStrip
        vehicleName={vehicleName}
        plateNumber={plateNumber}
        tankVolumeLiters={vehicle?.tank_volume_liters ?? 0}
        waybills={waybills}
        driversById={driversById}
        onDayClick={setSelected}
      />

      <VehicleDayFeed
        month={displayMonth}
        feed={feed}
        driversById={driversById}
        onGapClick={onGenerate ? focusGenerate : undefined}
        onWaybillClick={setSelected}
      />

      <WaybillDayDrawer
        open={selected != null}
        waybill={selected}
        vehicle={vehicle}
        periodWaybills={waybills}
        onClose={() => setSelected(null)}
        onSaved={(saved) => {
          const rechained = saved.rechained_draft_days ?? 0;
          setRechainedNote(rechained > 0 ? `Пересчитано дней: ${rechained}` : null);
        }}
      />
      <ManualWaybillDialog
        open={manualOpen}
        onClose={() => setManualOpen(false)}
        defaultVehicleId={vehicleId}
        defaultDate={periodFrom}
        periodWaybills={waybills}
      />
    </div>
  );
};
