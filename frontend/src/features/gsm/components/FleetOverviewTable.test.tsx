import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { FleetOverviewTable } from "@/features/gsm/components/FleetOverviewTable";
import type { FleetOverviewRow } from "@/features/gsm/types/gsm";

const row = (
  overrides: Partial<FleetOverviewRow> & { id: number; status: FleetOverviewRow["status"] },
): FleetOverviewRow => ({
  vehicle: { id: overrides.id, name: `Машина ${overrides.id}`, plate_number: `A${overrides.id}` },
  tx_count: 1,
  tx_liters: 10,
  tx_amount: 100,
  tx_last_date: "2026-08-10",
  wb_count: 1,
  wb_km: 100,
  wb_fuel_issued: 10,
  wb_last_date: "2026-08-10",
  red_days: 0,
  draft_count: 0,
  confirmed_count: 0,
  exported_count: 1,
  fuel_end_last: 20,
  liters_diff: 0,
  open_before: 0,
  open_before_month: null,
  chain_broken: false,
  status: overrides.status,
  ...overrides,
  vehicle: overrides.vehicle ?? {
    id: overrides.id,
    name: `Машина ${overrides.id}`,
    plate_number: `A${overrides.id}`,
  },
});

const renderTable = (
  rows: FleetOverviewRow[],
  extra?: Partial<Parameters<typeof FleetOverviewTable>[0]>,
) =>
  render(
    <FleetOverviewTable
      rows={rows}
      selectedIds={new Set()}
      onToggle={() => undefined}
      onToggleAllSelectable={() => undefined}
      expandedId={null}
      onToggleExpand={() => undefined}
      periodFrom="2026-08-01"
      {...extra}
    />,
  );

describe("FleetOverviewTable", () => {
  afterEach(() => {
    cleanup();
  });

  it("renders status labels from fleetStatusMeta", () => {
    renderTable([
      row({ id: 1, status: "needs_generation" }),
      row({ id: 2, status: "has_red_days" }),
      row({ id: 3, status: "ready" }),
    ]);
    expect(screen.getByText("Требуется генерация")).toBeInTheDocument();
    expect(screen.getByText("Есть красные дни")).toBeInTheDocument();
    expect(screen.getByText("Выгружено")).toBeInTheDocument();
  });

  it("hides liters badge at wb_count=0 and paints green/red otherwise", () => {
    renderTable([
      row({ id: 1, status: "needs_generation", wb_count: 0, liters_diff: 40 }),
      row({ id: 2, status: "ready", wb_count: 2, liters_diff: 0 }),
      row({ id: 3, status: "ready", wb_count: 2, liters_diff: -2 }),
    ]);
    expect(screen.queryByTestId("liters-diff-1")).not.toBeInTheDocument();
    expect(screen.getByTestId("liters-diff-2")).toHaveTextContent("0.0 л");
    expect(screen.getByTestId("liters-diff-3")).toHaveTextContent(/−2|-\s?2/);
  });

  it("selects a row and select-all skips no_data", () => {
    const selected = new Set<number>();
    const onToggle = vi.fn((id: number) => selected.add(id));
    const onToggleAll = vi.fn();
    const { rerender } = renderTable(
      [row({ id: 1, status: "ready" }), row({ id: 2, status: "no_data", tx_count: 0, wb_count: 0 })],
      { selectedIds: selected, onToggle, onToggleAllSelectable: onToggleAll },
    );
    fireEvent.click(screen.getByLabelText("Выбрать Машина 1"));
    expect(onToggle).toHaveBeenCalledWith(1);
    expect(screen.getByLabelText("Выбрать Машина 2")).toBeDisabled();

    fireEvent.click(screen.getByLabelText("Выбрать все"));
    expect(onToggleAll).toHaveBeenCalled();

    rerender(
      <FleetOverviewTable
        rows={[
          row({ id: 1, status: "ready" }),
          row({ id: 2, status: "no_data", tx_count: 0, wb_count: 0 }),
        ]}
        selectedIds={new Set([1])}
        onToggle={onToggle}
        onToggleAllSelectable={onToggleAll}
        expandedId={null}
        onToggleExpand={() => undefined}
        periodFrom="2026-08-01"
      />,
    );
    expect(screen.getByLabelText("Выбрать Машина 1")).toBeChecked();
    expect(screen.getByLabelText("Выбрать Машина 2")).not.toBeChecked();
  });

  it("shows the tail badge as «Июль не выгружен: N ПЛ»", () => {
    renderTable([
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
    ]);
    expect(screen.getByTestId("open-before-1")).toHaveTextContent(/Июль не выгружен: 6 ПЛ/);
    expect(screen.getByTestId("open-before-1")).not.toHaveTextContent(/до периода/i);
  });

  it("renders row Export as a button, not a span", () => {
    renderTable([row({ id: 1, status: "drafts_pending", exported_count: 0 })]);
    const exportBtn = screen.getByRole("button", { name: "Экспорт" });
    expect(exportBtn.tagName).toBe("BUTTON");
  });

  it("calls onExportKit when Export is clicked", () => {
    const onExportKit = vi.fn();
    const kitRow = row({ id: 1, status: "pending_export", exported_count: 0 });
    renderTable([kitRow], { onExportKit });
    fireEvent.click(screen.getByRole("button", { name: "Экспорт" }));
    expect(onExportKit).toHaveBeenCalledWith(kitRow);
  });

  it("hides Generate for a July-tail row on August and shows Export", () => {
    renderTable([
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
    ]);
    expect(screen.queryByRole("button", { name: "Сгенерировать" })).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Экспорт" })).toBeInTheDocument();
  });

  it("shows Generate when the period is the vehicle's tail month", () => {
    renderTable(
      [
        row({
          id: 1,
          status: "needs_generation",
          open_before: 6,
          open_before_month: "2026-07",
          wb_count: 0,
        }),
      ],
      { periodFrom: "2026-07-01" },
    );
    expect(screen.getByRole("button", { name: "Сгенерировать" })).toBeInTheDocument();
  });

  it("hides Export for chain_broken without a tail", () => {
    renderTable([
      row({
        id: 1,
        status: "drafts_pending",
        chain_broken: true,
        exported_count: 0,
        wb_count: 2,
      }),
    ]);
    expect(screen.queryByRole("button", { name: "Экспорт" })).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Пересчитать август/i })).toBeInTheDocument();
  });

  it("hides Recalc for July tail + chain_broken on August and shows Export only", () => {
    renderTable([
      row({
        id: 1,
        status: "needs_generation",
        chain_broken: true,
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 2,
      }),
    ]);
    expect(screen.queryByRole("button", { name: /Пересчитать/i })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Сгенерировать" })).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Экспорт" })).toBeInTheDocument();
  });

  it("shows «Пересчитать август» for chain_broken and calls onGenerate", () => {
    const onGenerate = vi.fn();
    const broken = row({
      id: 1,
      status: "needs_generation",
      chain_broken: true,
      wb_count: 2,
    });
    renderTable([broken], { onGenerate });
    const cta = screen.getByRole("button", { name: /Пересчитать август/i });
    expect(screen.queryByRole("button", { name: "Сгенерировать" })).not.toBeInTheDocument();
    fireEvent.click(cta);
    expect(onGenerate).toHaveBeenCalledWith(broken);
  });
});
