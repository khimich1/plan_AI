import { useMemo, useState, type CSSProperties } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { FleetOverviewTable, selectableVehicleIds } from "@/features/gsm/components/FleetOverviewTable";
import { VehicleGenerateDialog } from "@/features/gsm/components/VehicleGenerateDialog";
import { VehicleWaybillJournal } from "@/features/gsm/components/VehicleWaybillJournal";
import { formatGsmCodeMessage, formatGsmError } from "@/features/gsm/lib/gsmErrors";
import {
  currentMonthBounds,
  fleetOpenBeforeMonth,
  openBeforeMonthLabel,
  openBeforeSummary,
} from "@/features/gsm/lib/fleetStatus";
import { bulkKitConfirmMessages, planBulkGenerate, planKit } from "@/features/gsm/lib/exportGate";
import { monthBounds } from "@/features/gsm/lib/vehicleDayFeed";
import {
  useBulkGenerateMutation,
  useDownloadGsmUsageReportMutation,
  useGsmOverviewQuery,
} from "@/features/gsm/hooks/useGsmQueries";
import type { BulkGenerateVehicleResult, FleetOverviewRow } from "@/features/gsm/types/gsm";

const sectionStyle: CSSProperties = { display: "grid", gap: "1rem" };

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

const barStyle: CSSProperties = {
  display: "flex",
  gap: "0.75rem",
  flexWrap: "wrap",
  alignItems: "center",
};

type Props = {
  onGenerateVehicle?: (row: FleetOverviewRow) => void;
};

type KitConfirm = {
  messages: string[];
  vehicleIds: number[];
  from: string;
  to: string;
};

const vehicleLabel = (row: FleetOverviewRow | undefined, vehicleId: number): string =>
  row ? `${row.vehicle.name} (${row.vehicle.plate_number})` : `Машина ${vehicleId}`;

const skipResults = (skipped: ReturnType<typeof planBulkGenerate>["skipped"]): BulkGenerateVehicleResult[] =>
  skipped.map((item) => ({
    vehicle_id: item.vehicleId,
    ok: false,
    error: { code: "gsm_kit_gate", message: item.reason },
  }));

export const FleetOverviewView = ({ onGenerateVehicle }: Props) => {
  const defaults = currentMonthBounds();
  const [periodFrom, setPeriodFrom] = useState(defaults.from);
  const [periodTo, setPeriodTo] = useState(defaults.to);
  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [expandedId, setExpandedId] = useState<number | null>(null);
  const [generateRow, setGenerateRow] = useState<FleetOverviewRow | null>(null);
  const [generatePeriod, setGeneratePeriod] = useState<{ from: string; to: string } | null>(null);
  const [bulkReport, setBulkReport] = useState<BulkGenerateVehicleResult[] | null>(null);
  const [bulkError, setBulkError] = useState<string | null>(null);
  const [kitExclusions, setKitExclusions] = useState<ReturnType<typeof planKit>["excluded"]>([]);
  const [kitConfirm, setKitConfirm] = useState<KitConfirm | null>(null);
  const [usageReportInfo, setUsageReportInfo] = useState<string | null>(null);

  const overviewQuery = useGsmOverviewQuery({ periodFrom, periodTo });
  const bulkGenerate = useBulkGenerateMutation();
  const usageReport = useDownloadGsmUsageReportMutation();
  const rows = overviewQuery.data ?? [];
  const rowsById = useMemo(() => new Map(rows.map((row) => [row.vehicle.id, row])), [rows]);

  const { pl: openBeforeTotal, vehicles: openBeforeVehicles } = useMemo(
    () => openBeforeSummary(rows),
    [rows],
  );
  const tailYm = useMemo(() => fleetOpenBeforeMonth(rows), [rows]);
  const tailLabel = tailYm ? openBeforeMonthLabel(tailYm) : null;

  const toggle = (vehicleId: number) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(vehicleId)) next.delete(vehicleId);
      else next.add(vehicleId);
      return next;
    });
  };

  const toggleAllSelectable = () => {
    const selectable = selectableVehicleIds(rows);
    setSelectedIds((prev) => {
      const allOn = selectable.length > 0 && selectable.every((id) => prev.has(id));
      return allOn ? new Set() : new Set(selectable);
    });
  };

  const selectedList = [...selectedIds].sort((a, b) => a - b);
  const bulkBusy = bulkGenerate.isPending || usageReport.isPending;

  const openGenerate = (row: FleetOverviewRow, bounds?: { from: string; to: string }) => {
    if (onGenerateVehicle) {
      onGenerateVehicle(row);
      return;
    }
    setGeneratePeriod(bounds ?? null);
    setGenerateRow(row);
  };

  const handleBulkGenerate = async () => {
    setBulkError(null);
    setBulkReport(null);
    if (selectedList.length === 0) {
      return;
    }
    const plan = planBulkGenerate(rows, selectedList, periodFrom);
    const localSkips = skipResults(plan.skipped);
    if (plan.eligibleIds.length === 0) {
      setBulkReport(localSkips);
      return;
    }
    try {
      const result = await bulkGenerate.mutateAsync({
        vehicle_ids: plan.eligibleIds,
        period_from: periodFrom,
        period_to: periodTo,
      });
      setBulkReport(
        [...localSkips, ...result.results].sort((a, b) => a.vehicle_id - b.vehicle_id),
      );
    } catch (err) {
      setBulkError(formatGsmError(err));
    }
  };

  const runKit = async (vehicleIds: number[], from: string, to: string) => {
    setBulkError(null);
    try {
      await usageReport.mutateAsync({
        period_from: from,
        period_to: to,
        vehicle_ids: vehicleIds,
      });
      setUsageReportInfo("Скачан zip с отчётом об использовании ГСМ.");
    } catch (err) {
      setBulkError(formatGsmError(err));
    }
  };

  const startKit = (plan: ReturnType<typeof planKit>, from: string, to: string) => {
    setKitExclusions(plan.excluded);
    if (plan.cleanIds.length === 0) {
      return;
    }
    const confirms = bulkKitConfirmMessages(plan);
    if (confirms.length > 0) {
      setKitConfirm({ messages: confirms, vehicleIds: plan.cleanIds, from, to });
      return;
    }
    void runKit(plan.cleanIds, from, to);
  };

  const handlePeriodReport = () => {
    setUsageReportInfo(null);
    setBulkError(null);
    const selected = selectedList.length > 0 ? selectedList : null;
    startKit(planKit(rows, selected, periodFrom), periodFrom, periodTo);
  };

  const handleExportKit = (row: FleetOverviewRow) => {
    setUsageReportInfo(null);
    setBulkError(null);
    if (row.open_before > 0 && row.open_before_month) {
      const bounds = monthBounds(row.open_before_month);
      setKitExclusions([]);
      void runKit([row.vehicle.id], bounds.from, bounds.to);
      return;
    }
    startKit(planKit([row], [row.vehicle.id], periodFrom), periodFrom, periodTo);
  };

  const handleJournalGenerate = (row: FleetOverviewRow, bounds: { from: string; to: string }) => {
    const currentYm = periodFrom.slice(0, 7);
    const isTailMonth = row.open_before_month != null && currentYm === row.open_before_month;
    if (row.open_before > 0 && !isTailMonth) {
      const tailMonth = openBeforeMonthLabel(row.open_before_month ?? currentYm);
      setBulkError(`${vehicleLabel(row, row.vehicle.id)}: сначала выгрузите ${tailMonth}`);
      return;
    }
    openGenerate(row, bounds);
  };

  if (overviewQuery.isLoading) {
    return (
      <section style={sectionStyle} aria-label="Обзор флота ГСМ">
        <Spinner />
      </section>
    );
  }

  if (overviewQuery.error) {
    return (
      <section style={sectionStyle} aria-label="Обзор флота ГСМ">
        <Alert tone="error">{formatGsmError(overviewQuery.error)}</Alert>
      </section>
    );
  }

  return (
    <section style={sectionStyle} aria-label="Обзор флота ГСМ">
      <header>
        <h2 style={{ margin: 0, fontSize: "1.25rem" }}>Обзор флота</h2>
        <p style={{ margin: "0.35rem 0 0", color: "#475467" }}>
          Статусы машин за период, журналы ПЛ и массовые действия.
        </p>
      </header>

      {openBeforeTotal > 0 && tailLabel && tailYm && (
        <button
          type="button"
          data-testid="open-before-banner"
          onClick={() => {
            const bounds = monthBounds(tailYm);
            setPeriodFrom(bounds.from);
            setPeriodTo(bounds.to);
          }}
          style={{
            textAlign: "left",
            border: "1px solid #fec84b",
            background: "#fffaeb",
            borderRadius: 12,
            padding: "0.75rem 1rem",
            cursor: "pointer",
            color: "#93370d",
          }}
        >
          {tailLabel} не выгружен: {openBeforeTotal} ПЛ по {openBeforeVehicles} машинам. Открыть{" "}
          {tailLabel}.
        </button>
      )}

      <div style={controlsStyle}>
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
        <div style={{ display: "grid", gap: 4 }}>
          <Button type="button" variant="secondary" onClick={handlePeriodReport} disabled={bulkBusy}>
            Отчёт за период
          </Button>
          <span style={{ fontSize: "0.8rem", color: "#667085" }}>сводка и путевые</span>
        </div>
      </div>

      <div style={barStyle}>
        <Button
          type="button"
          disabled={selectedList.length === 0 || bulkBusy}
          onClick={() => void handleBulkGenerate()}
        >
          Сгенерировать выбранные
        </Button>
        <span style={{ color: "#667085", fontSize: "0.85rem" }}>
          Выбрано: {selectedList.length}
        </span>
      </div>

      {bulkError && <Alert tone="error">{bulkError}</Alert>}
      {usageReportInfo && <Alert tone="success">{usageReportInfo}</Alert>}
      {kitExclusions.length > 0 && (
        <Alert tone="warning">
          <div data-testid="kit-exclusions">
            Исключены из комплекта:
            <ul style={{ margin: "0.4rem 0 0", paddingLeft: "1.2rem" }}>
              {kitExclusions.map((item) => (
                <li key={item.vehicleId}>
                  {item.label}: {item.reason}
                </li>
              ))}
            </ul>
          </div>
        </Alert>
      )}
      {bulkReport && (
        <Alert tone="info">
          <ul data-testid="bulk-generate-report" style={{ margin: 0, paddingLeft: "1.2rem" }}>
            {bulkReport.map((item) => {
              const row = rowsById.get(item.vehicle_id);
              const label = vehicleLabel(row, item.vehicle_id);
              if (item.ok) {
                return (
                  <li key={item.vehicle_id} data-testid={`bulk-result-${item.vehicle_id}`}>
                    {label}: готово
                  </li>
                );
              }
              const message = formatGsmCodeMessage(item.error?.code, item.error?.message ?? "ошибка");
              return (
                <li key={item.vehicle_id} data-testid={`bulk-result-${item.vehicle_id}`}>
                  {label}: {message}
                  {item.error?.code === "gsm_start_required" && row && (
                    <>
                      {" "}
                      <Button
                        type="button"
                        variant="secondary"
                        onClick={() => openGenerate(row)}
                      >
                        Указать старт
                      </Button>
                    </>
                  )}
                </li>
              );
            })}
          </ul>
        </Alert>
      )}

      <FleetOverviewTable
        rows={rows}
        selectedIds={selectedIds}
        onToggle={toggle}
        onToggleAllSelectable={toggleAllSelectable}
        expandedId={expandedId}
        onToggleExpand={(id) => setExpandedId((cur) => (cur === id ? null : id))}
        periodFrom={periodFrom}
        onGenerate={onGenerateVehicle ?? ((row) => openGenerate(row))}
        onExportKit={handleExportKit}
        renderExpanded={(row) => (
          <VehicleWaybillJournal
            vehicleId={row.vehicle.id}
            vehicleName={row.vehicle.name}
            plateNumber={row.vehicle.plate_number}
            periodFrom={periodFrom}
            periodTo={periodTo}
            onPeriodChange={(bounds) => {
              setPeriodFrom(bounds.from);
              setPeriodTo(bounds.to);
            }}
            onGenerate={(bounds) => handleJournalGenerate(row, bounds)}
          />
        )}
      />
      <VehicleGenerateDialog
        open={generateRow != null}
        row={generateRow}
        periodFrom={generatePeriod?.from ?? periodFrom}
        periodTo={generatePeriod?.to ?? periodTo}
        onClose={() => {
          setGenerateRow(null);
          setGeneratePeriod(null);
        }}
      />
      <Modal
        open={kitConfirm != null}
        onClose={() => setKitConfirm(null)}
        title="Отчёт за период"
        maxWidth={480}
      >
        <div style={{ display: "grid", gap: "0.85rem" }}>
          {kitConfirm?.messages.map((message) => (
            <p key={message} style={{ margin: 0, color: "#475467" }}>
              {message}
            </p>
          ))}
          <div style={{ display: "flex", gap: "0.75rem", justifyContent: "flex-end" }}>
            <Button type="button" variant="secondary" onClick={() => setKitConfirm(null)}>
              Отмена
            </Button>
            <Button
              type="button"
              disabled={usageReport.isPending}
              onClick={() => {
                const pending = kitConfirm;
                setKitConfirm(null);
                if (!pending) return;
                void runKit(pending.vehicleIds, pending.from, pending.to);
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
