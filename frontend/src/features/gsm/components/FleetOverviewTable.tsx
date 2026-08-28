import type { CSSProperties, ReactNode } from "react";
import type { FleetOverviewRow } from "@/features/gsm/types/gsm";
import {
  TONE_COLORS,
  fleetStatusMeta,
  litersDiffHidden,
  litersDiffOk,
  openBeforeMonthLabel,
} from "@/features/gsm/lib/fleetStatus";
import { formatKm, formatLiters } from "@/features/gsm/lib/waybillWarnings";
import { Button } from "@/shared/ui/Button";

type Props = {
  rows: FleetOverviewRow[];
  selectedIds: ReadonlySet<number>;
  onToggle: (vehicleId: number) => void;
  onToggleAllSelectable: () => void;
  expandedId: number | null;
  onToggleExpand: (vehicleId: number) => void;
  periodFrom: string;
  onGenerate?: (row: FleetOverviewRow) => void;
  onExportKit?: (row: FleetOverviewRow) => void;
  renderExpanded?: (row: FleetOverviewRow) => ReactNode;
};

const tableWrap: CSSProperties = {
  overflowX: "auto",
  border: "1px solid #eaecf0",
  borderRadius: 14,
  background: "#ffffff",
};

const th: CSSProperties = {
  textAlign: "left",
  padding: "0.65rem 0.75rem",
  fontSize: "0.8rem",
  color: "#475467",
  borderBottom: "1px solid #eaecf0",
  whiteSpace: "nowrap",
};

const td: CSSProperties = {
  padding: "0.65rem 0.75rem",
  fontSize: "0.9rem",
  borderBottom: "1px solid #f2f4f7",
  verticalAlign: "middle",
};

const formatDiff = (diff: number): string => {
  const sign = diff > 0 ? "+" : "";
  return `${sign}${diff.toLocaleString("ru-RU", { maximumFractionDigits: 2 })} л`;
};

export const selectableVehicleIds = (rows: FleetOverviewRow[]): number[] =>
  rows.filter((row) => row.status !== "no_data").map((row) => row.vehicle.id);

export const FleetOverviewTable = ({
  rows,
  selectedIds,
  onToggle,
  onToggleAllSelectable,
  expandedId,
  onToggleExpand,
  periodFrom,
  onGenerate,
  onExportKit,
  renderExpanded,
}: Props) => {
  const selectable = selectableVehicleIds(rows);
  const allSelected = selectable.length > 0 && selectable.every((id) => selectedIds.has(id));

  return (
    <div style={tableWrap}>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead>
          <tr>
            <th style={th}>
              <input
                type="checkbox"
                aria-label="Выбрать все"
                checked={allSelected}
                onChange={onToggleAllSelectable}
              />
            </th>
            <th style={th}>Машина</th>
            <th style={th}>Статус</th>
            <th style={th}>Транзакции</th>
            <th style={th}>ПЛ</th>
            <th style={th}>Красные дни</th>
            <th style={th}>Бак</th>
            <th style={th}>Расхождение</th>
            <th style={th} />
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const meta = fleetStatusMeta(row.status);
            const tone = TONE_COLORS[meta.tone];
            const expanded = expandedId === row.vehicle.id;
            const selectableRow = row.status !== "no_data";
            const hideDiff = litersDiffHidden(row.wb_count);
            const diffOk = litersDiffOk(row.liters_diff);
            return (
              <RowGroup
                key={row.vehicle.id}
                row={row}
                metaLabel={meta.label}
                tone={tone}
                expanded={expanded}
                selectableRow={selectableRow}
                selected={selectedIds.has(row.vehicle.id)}
                hideDiff={hideDiff}
                diffOk={diffOk}
                onToggle={onToggle}
                onToggleExpand={onToggleExpand}
                periodFrom={periodFrom}
                onGenerate={onGenerate}
                onExportKit={onExportKit}
                renderExpanded={renderExpanded}
              />
            );
          })}
        </tbody>
      </table>
    </div>
  );
};

type RowGroupProps = {
  row: FleetOverviewRow;
  metaLabel: string;
  tone: { bg: string; fg: string };
  expanded: boolean;
  selectableRow: boolean;
  selected: boolean;
  hideDiff: boolean;
  diffOk: boolean;
  onToggle: (vehicleId: number) => void;
  onToggleExpand: (vehicleId: number) => void;
  periodFrom: string;
  onGenerate?: (row: FleetOverviewRow) => void;
  onExportKit?: (row: FleetOverviewRow) => void;
  renderExpanded?: (row: FleetOverviewRow) => ReactNode;
};

const RowGroup = ({
  row,
  metaLabel,
  tone,
  expanded,
  selectableRow,
  selected,
  hideDiff,
  diffOk,
  onToggle,
  onToggleExpand,
  periodFrom,
  onGenerate,
  onExportKit,
  renderExpanded,
}: RowGroupProps) => (
  <>
    <tr style={{ background: row.status === "no_data" ? "#f9fafb" : undefined }}>
      <td style={td}>
        <input
          type="checkbox"
          aria-label={`Выбрать ${row.vehicle.name}`}
          checked={selected}
          disabled={!selectableRow}
          onChange={() => onToggle(row.vehicle.id)}
        />
      </td>
      <td style={td}>
        <button
          type="button"
          onClick={() => onToggleExpand(row.vehicle.id)}
          aria-expanded={expanded}
          style={{
            border: "none",
            background: "transparent",
            cursor: "pointer",
            textAlign: "left",
            padding: 0,
            fontWeight: 600,
            color: "#101828",
          }}
        >
          {row.vehicle.name}
        </button>
        <div style={{ color: "#667085", fontSize: "0.8rem" }}>{row.vehicle.plate_number}</div>
        {row.open_before > 0 && (
          <span
            data-testid={`open-before-${row.vehicle.id}`}
            style={{
              display: "inline-block",
              marginTop: 4,
              fontSize: "0.75rem",
              color: "#93370d",
              background: "#fef0c7",
              borderRadius: 8,
              padding: "0.1rem 0.4rem",
            }}
          >
            {row.open_before_month
              ? `${openBeforeMonthLabel(row.open_before_month)} не выгружен: ${row.open_before} ПЛ`
              : `не выгружено: ${row.open_before} ПЛ`}
          </span>
        )}
      </td>
      <td style={td}>
        <span
          style={{
            display: "inline-block",
            borderRadius: 999,
            padding: "0.15rem 0.55rem",
            background: tone.bg,
            color: tone.fg,
            fontWeight: 600,
            fontSize: "0.8rem",
          }}
        >
          {metaLabel}
        </span>
      </td>
      <td style={td}>
        {row.tx_count} · {formatLiters(row.tx_liters)}
      </td>
      <td style={td}>
        {row.wb_count} дн. · {formatKm(row.wb_km)}
      </td>
      <td style={td}>{row.red_days}</td>
      <td style={td}>{formatLiters(row.fuel_end_last)}</td>
      <td style={td}>
        {!hideDiff && (
          <span
            data-testid={`liters-diff-${row.vehicle.id}`}
            style={{
              fontWeight: 600,
              color: diffOk ? "#067647" : "#b42318",
            }}
          >
            {diffOk ? "0.0 л" : formatDiff(row.liters_diff)}
          </span>
        )}
      </td>
      <td style={td}>
        <RowActions
          row={row}
          periodFrom={periodFrom}
          onGenerate={onGenerate}
          onExportKit={onExportKit}
        />
      </td>
    </tr>
    {expanded && (
      <tr>
        <td style={{ ...td, background: "#f8fafc" }} colSpan={9}>
          {renderExpanded?.(row)}
        </td>
      </tr>
    )}
  </>
);

const actionsStyle: CSSProperties = {
  display: "flex",
  gap: "0.35rem",
  flexWrap: "wrap",
  alignItems: "center",
};

type RowActionsProps = {
  row: FleetOverviewRow;
  periodFrom: string;
  onGenerate?: (row: FleetOverviewRow) => void;
  onExportKit?: (row: FleetOverviewRow) => void;
};

const RowActions = ({ row, periodFrom, onGenerate, onExportKit }: RowActionsProps) => {
  const periodYm = periodFrom.slice(0, 7);
  const periodLabel = periodYm.length >= 7 ? openBeforeMonthLabel(periodYm) : "";
  const isTailPeriod =
    row.open_before_month != null && periodYm === row.open_before_month;
  const blockedByTail = row.open_before > 0 && !isTailPeriod;
  const showRecalc = row.chain_broken && !blockedByTail;
  const showGenerate = !row.chain_broken && row.status === "needs_generation" && !blockedByTail;
  const showExport =
    row.open_before > 0 ||
    (!row.chain_broken && (row.status === "drafts_pending" || row.status === "pending_export"));

  return (
    <div style={actionsStyle}>
      {showRecalc && (
        <Button type="button" onClick={() => onGenerate?.(row)}>
          Пересчитать {periodLabel}
        </Button>
      )}
      {showGenerate && (
        <Button type="button" onClick={() => onGenerate?.(row)}>
          Сгенерировать
        </Button>
      )}
      {showExport && (
        <Button type="button" variant="secondary" onClick={() => onExportKit?.(row)}>
          Экспорт
        </Button>
      )}
    </div>
  );
};

