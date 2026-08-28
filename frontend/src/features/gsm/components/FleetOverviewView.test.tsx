import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { FleetOverviewView } from "@/features/gsm/components/FleetOverviewView";
import { currentMonthBounds } from "@/features/gsm/lib/fleetStatus";
import { monthBounds, shiftMonth } from "@/features/gsm/lib/vehicleDayFeed";
import type { FleetOverviewRow } from "@/features/gsm/types/gsm";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";

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

const { overviewState, bulkGenerateAsync, exportAsync, usageReportAsync } = vi.hoisted(() => ({
  overviewState: { rows: [] as FleetOverviewRow[] },
  bulkGenerateAsync: vi.fn(),
  exportAsync: vi.fn(),
  usageReportAsync: vi.fn(),
}));

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmOverviewQuery: () => ({
    isLoading: false,
    error: null,
    data: overviewState.rows,
  }),
  useGsmWaybillsQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGsmTransactionsQuery: () => ({
    isLoading: false,
    error: null,
    data: { rows: [], total_count: 0, sum_liters: 0, sum_amount: 0 },
  }),
  useGsmDriversQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGsmVehiclesQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGenerateGsmWaybillsMutation: () => ({ mutateAsync: vi.fn(), isPending: false }),
  useBulkGenerateMutation: () => ({ mutateAsync: bulkGenerateAsync, isPending: false }),
  useExportGsmWaybillsMutation: () => ({ mutateAsync: exportAsync, isPending: false }),
  useDownloadGsmUsageReportMutation: () => ({ mutateAsync: usageReportAsync, isPending: false }),
  useGsmRoutesQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGsmStationsQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGsmSettingsQuery: () => ({ isLoading: false, error: null, data: {} }),
  usePatchGsmWaybillMutation: () => ({ mutateAsync: vi.fn(), isPending: false }),
  useCreateGsmWaybillMutation: () => ({ mutateAsync: vi.fn(), isPending: false }),
}));

function renderOverview() {
  const qc = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  return render(
    <QueryClientProvider client={qc}>
      <FleetOverviewView />
    </QueryClientProvider>,
  );
}

const setAugustPeriod = () => {
  fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2026-08-01" } });
  fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2026-08-31" } });
};

describe("FleetOverviewView tails banner", () => {
  afterEach(() => {
    cleanup();
  });

  it("shows the July tail banner and opens that month, not the previous calendar month", () => {
    overviewState.rows = [
      row({ id: 1, status: "ready", open_before: 2, open_before_month: "2026-07" }),
      row({ id: 2, status: "no_data", tx_count: 0, wb_count: 0, open_before: 0 }),
    ];
    renderOverview();
    const banner = screen.getByTestId("open-before-banner");
    expect(banner).toHaveTextContent(/Июль не выгружен: 2 ПЛ по 1 машин/);
    expect(banner).toHaveTextContent(/Открыть Июль/);
    expect(banner).not.toHaveTextContent(/до периода/i);
    expect(screen.getByTestId("open-before-1")).toBeInTheDocument();
    expect(screen.queryByTestId("open-before-2")).not.toBeInTheDocument();

    fireEvent.click(banner);
    expect(screen.getByLabelText("Период с")).toHaveValue("2026-07-01");
    expect(screen.getByLabelText("Период по")).toHaveValue("2026-07-31");
  });
});

describe("FleetOverviewView bulk bar", () => {
  afterEach(() => {
    cleanup();
    bulkGenerateAsync.mockReset();
    exportAsync.mockReset();
    usageReportAsync.mockReset();
  });

  it("does not show the zip-PL bulk export button", () => {
    overviewState.rows = [row({ id: 1, status: "ready" })];
    renderOverview();
    expect(screen.queryByRole("button", { name: /Экспорт zip выбранных/ })).not.toBeInTheDocument();
    expect(screen.queryByText("Экспорт zip выбранных")).not.toBeInTheDocument();
  });

  it("does not call bulk generate when nothing is selected", () => {
    overviewState.rows = [row({ id: 1, status: "needs_generation", wb_count: 0 })];
    renderOverview();
    const btn = screen.getByRole("button", { name: "Сгенерировать выбранные" });
    expect(btn).toBeDisabled();
    fireEvent.click(btn);
    expect(bulkGenerateAsync).not.toHaveBeenCalled();
  });

  it("shows a per-vehicle error in the report while others succeed", async () => {
    overviewState.rows = [
      row({ id: 1, status: "ready" }),
      row({ id: 3, status: "needs_generation", wb_count: 0 }),
    ];
    bulkGenerateAsync.mockResolvedValue({
      results: [
        {
          vehicle_id: 1,
          ok: true,
          result: {
            days_created: 5,
            waybills: [],
            warnings: [],
            problematic_days: [],
            manual_days: 0,
          },
        },
        {
          vehicle_id: 3,
          ok: false,
          error: { code: "gsm_routes_required", message: "vehicle has no routes in library" },
        },
      ],
    });
    renderOverview();
    fireEvent.click(screen.getByLabelText("Выбрать Машина 1"));
    fireEvent.click(screen.getByLabelText("Выбрать Машина 3"));
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать выбранные" }));
    expect(await screen.findByTestId("bulk-result-1")).toHaveTextContent("готово");
    expect(screen.getByTestId("bulk-result-3")).toHaveTextContent(
      "У машины нет маршрутов в библиотеке.",
    );
  });

  it("skips a July-tail vehicle in bulk generate and only POSTs eligible ids", async () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
      row({ id: 2, status: "needs_generation", wb_count: 0 }),
    ];
    bulkGenerateAsync.mockResolvedValue({
      results: [
        {
          vehicle_id: 2,
          ok: true,
          result: {
            days_created: 3,
            waybills: [],
            warnings: [],
            problematic_days: [],
            manual_days: 0,
          },
        },
      ],
    });
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByLabelText("Выбрать Машина 1"));
    fireEvent.click(screen.getByLabelText("Выбрать Машина 2"));
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать выбранные" }));
    expect(bulkGenerateAsync).toHaveBeenCalledWith({
      vehicle_ids: [2],
      period_from: "2026-08-01",
      period_to: "2026-08-31",
    });
    expect(await screen.findByTestId("bulk-result-1")).toHaveTextContent(/сначала выгрузите Июль/);
    expect(screen.getByTestId("bulk-result-2")).toHaveTextContent("готово");
  });

  it("skips a chain_broken vehicle in bulk generate and only POSTs eligible ids", async () => {
    overviewState.rows = [
      row({ id: 1, status: "needs_generation", chain_broken: true, wb_count: 2 }),
      row({ id: 2, status: "needs_generation", wb_count: 0 }),
    ];
    bulkGenerateAsync.mockResolvedValue({
      results: [
        {
          vehicle_id: 2,
          ok: true,
          result: {
            days_created: 3,
            waybills: [],
            warnings: [],
            problematic_days: [],
            manual_days: 0,
          },
        },
      ],
    });
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByLabelText("Выбрать Машина 1"));
    fireEvent.click(screen.getByLabelText("Выбрать Машина 2"));
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать выбранные" }));
    expect(bulkGenerateAsync).toHaveBeenCalledWith({
      vehicle_ids: [2],
      period_from: "2026-08-01",
      period_to: "2026-08-31",
    });
    expect(await screen.findByTestId("bulk-result-1")).toHaveTextContent(/пересчитайте Август/);
    expect(screen.getByTestId("bulk-result-2")).toHaveTextContent("готово");
  });

  it("does not POST bulk generate when every selected vehicle is skipped", async () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
      row({ id: 2, status: "needs_generation", chain_broken: true, wb_count: 2 }),
    ];
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByLabelText("Выбрать Машина 1"));
    fireEvent.click(screen.getByLabelText("Выбрать Машина 2"));
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать выбранные" }));
    expect(bulkGenerateAsync).not.toHaveBeenCalled();
    expect(await screen.findByTestId("bulk-result-1")).toHaveTextContent(/сначала выгрузите Июль/);
    expect(screen.getByTestId("bulk-result-2")).toHaveTextContent(/пересчитайте Август/);
  });

  it("opens the generate dialog when bulk reports gsm_start_required", async () => {
    overviewState.rows = [row({ id: 4, status: "needs_generation", wb_count: 0 })];
    bulkGenerateAsync.mockResolvedValue({
      results: [
        {
          vehicle_id: 4,
          ok: false,
          error: {
            code: "gsm_start_required",
            message: "fuel_start and odometer_start required when no confirmed waybill exists",
          },
        },
      ],
    });
    renderOverview();
    fireEvent.click(screen.getByLabelText("Выбрать Машина 4"));
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать выбранные" }));
    fireEvent.click(await screen.findByRole("button", { name: "Указать старт" }));
    expect(screen.getByText("Генерация: Машина 4")).toBeInTheDocument();
  });
});

describe("FleetOverviewView usage report", () => {
  afterEach(() => {
    cleanup();
    usageReportAsync.mockReset();
  });

  it("hints that the period report is summary plus waybills", () => {
    overviewState.rows = [row({ id: 1, status: "ready", exported_count: 0, wb_count: 1 })];
    renderOverview();
    expect(screen.getByText("сводка и путевые")).toBeInTheDocument();
  });

  it("sends planKit clean ids when nobody is selected", async () => {
    overviewState.rows = [row({ id: 1, status: "drafts_pending", exported_count: 0, wb_count: 1 })];
    usageReportAsync.mockResolvedValue({
      blob: new Blob(),
      filename: "gsm_usage_report.zip",
    });
    renderOverview();
    const bounds = currentMonthBounds();
    fireEvent.click(screen.getByRole("button", { name: "Отчёт за период" }));
    expect(usageReportAsync).toHaveBeenCalledWith({
      period_from: bounds.from,
      period_to: bounds.to,
      vehicle_ids: [1],
    });
    expect(await screen.findByText("Скачан zip с отчётом об использовании ГСМ.")).toBeInTheDocument();
  });

  it("excludes a July tail from an August report and POSTs only clean ids, never null", async () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        exported_count: 0,
        wb_count: 0,
      }),
      row({ id: 2, status: "drafts_pending", exported_count: 0, wb_count: 4 }),
    ];
    usageReportAsync.mockResolvedValue({
      blob: new Blob(),
      filename: "gsm_usage_report.zip",
    });
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByRole("button", { name: "Отчёт за период" }));

    expect(await screen.findByTestId("kit-exclusions")).toHaveTextContent(
      /Машина 1 \(A1\): Сначала выгрузите Июль/,
    );
    expect(usageReportAsync).toHaveBeenCalledTimes(1);
    expect(usageReportAsync).toHaveBeenCalledWith({
      period_from: "2026-08-01",
      period_to: "2026-08-31",
      vehicle_ids: [2],
    });
    expect(usageReportAsync.mock.calls[0][0].vehicle_ids).not.toBeNull();
  });

  it("does not POST the usage report when every vehicle is excluded", () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
    ];
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByRole("button", { name: "Отчёт за период" }));
    expect(screen.getByTestId("kit-exclusions")).toHaveTextContent(/Сначала выгрузите Июль/);
    expect(usageReportAsync).not.toHaveBeenCalled();
  });

  it("excludes chain_broken from an August report and POSTs only the clean neighbor", async () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "drafts_pending",
        chain_broken: true,
        exported_count: 0,
        wb_count: 2,
      }),
      row({ id: 2, status: "drafts_pending", exported_count: 0, wb_count: 4 }),
    ];
    usageReportAsync.mockResolvedValue({
      blob: new Blob(),
      filename: "gsm_usage_report.zip",
    });
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByRole("button", { name: "Отчёт за период" }));

    expect(await screen.findByTestId("kit-exclusions")).toHaveTextContent(
      /Машина 1 \(A1\): Пересчитайте Август: бак не сходится с предыдущим/,
    );
    expect(usageReportAsync).toHaveBeenCalledTimes(1);
    expect(usageReportAsync).toHaveBeenCalledWith({
      period_from: "2026-08-01",
      period_to: "2026-08-31",
      vehicle_ids: [2],
    });
  });

  it("passes selected overview vehicle ids through planKit into the usage report", async () => {
    overviewState.rows = [
      row({ id: 1, status: "drafts_pending", exported_count: 0 }),
      row({ id: 2, status: "drafts_pending", exported_count: 0 }),
    ];
    usageReportAsync.mockResolvedValue({
      blob: new Blob(),
      filename: "gsm_usage_report.zip",
    });
    renderOverview();

    fireEvent.click(screen.getByLabelText("Выбрать Машина 1"));
    fireEvent.click(screen.getByRole("button", { name: "Отчёт за период" }));

    const bounds = currentMonthBounds();
    expect(usageReportAsync).toHaveBeenCalledWith({
      period_from: bounds.from,
      period_to: bounds.to,
      vehicle_ids: [1],
    });
  });
});

describe("FleetOverviewView row actions", () => {
  afterEach(() => {
    cleanup();
    usageReportAsync.mockReset();
  });

  it("opens VehicleGenerateDialog from «Пересчитать август» with force off", () => {
    overviewState.rows = [
      row({ id: 1, status: "needs_generation", chain_broken: true, wb_count: 2 }),
    ];
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByRole("button", { name: /Пересчитать август/i }));

    const dialog = screen.getByRole("dialog");
    expect(within(dialog).getByText("Генерация: Машина 1")).toBeInTheDocument();
    expect(within(dialog).getByLabelText("Период с")).toHaveValue("2026-08-01");
    expect(within(dialog).getByLabelText("Период по")).toHaveValue("2026-08-31");
    expect(within(dialog).getByLabelText("Перезаписать confirmed")).not.toBeChecked();
  });

  it("exports the current period for a row without a tail", async () => {
    overviewState.rows = [
      row({ id: 1, status: "drafts_pending", exported_count: 0, wb_count: 4 }),
    ];
    usageReportAsync.mockResolvedValue({
      blob: new Blob(),
      filename: "gsm_usage_report.zip",
    });
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByRole("button", { name: "Экспорт" }));
    expect(usageReportAsync).toHaveBeenCalledWith({
      period_from: "2026-08-01",
      period_to: "2026-08-31",
      vehicle_ids: [1],
    });
  });

  it("exports the tail month for a row with open_before", async () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
    ];
    usageReportAsync.mockResolvedValue({
      blob: new Blob(),
      filename: "gsm_usage_report.zip",
    });
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByRole("button", { name: "Экспорт" }));
    expect(usageReportAsync).toHaveBeenCalledWith({
      period_from: "2026-07-01",
      period_to: "2026-07-31",
      vehicle_ids: [1],
    });
  });
});

describe("FleetOverviewView one-month clock", () => {
  afterEach(() => {
    cleanup();
  });

  it("shows August calendar for 01.08–31.08 filters and July in inputs and grid after previous month", () => {
    overviewState.rows = [row({ id: 4, status: "has_red_days" })];
    renderOverview();

    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2026-08-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2026-08-31" } });
    fireEvent.click(screen.getByRole("button", { name: "Машина 4" }));

    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Август 2026 г.");
    expect(screen.getByTestId("cal-month-label")).not.toHaveTextContent(/июл/i);
    expect(screen.getByTestId("cal-day-2026-08-01")).toBeInTheDocument();
    expect(screen.getByTestId("cal-day-2026-08-31")).toBeInTheDocument();
    expect(screen.queryByTestId("cal-day-2026-07-31")).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Предыдущий месяц" }));

    expect(screen.getByLabelText("Период с")).toHaveValue("2026-07-01");
    expect(screen.getByLabelText("Период по")).toHaveValue("2026-07-31");
    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Июль 2026 г.");
    expect(screen.getByTestId("cal-day-2026-07-01")).toBeInTheDocument();
    expect(screen.getByTestId("cal-day-2026-07-31")).toBeInTheDocument();
    expect(screen.queryByTestId("cal-day-2026-08-01")).not.toBeInTheDocument();
  });
});

describe("FleetOverviewView journal generate gate", () => {
  afterEach(() => {
    cleanup();
  });

  it("blocks journal generate for a July tail on August without opening the dialog", () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
    ];
    renderOverview();
    setAugustPeriod();
    fireEvent.click(screen.getByRole("button", { name: "Машина 1" }));
    fireEvent.click(screen.getByTestId("journal-generate-btn"));
    expect(screen.getByText(/Машина 1 \(A1\): сначала выгрузите Июль/)).toBeInTheDocument();
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it("opens generate from the journal when the period is the vehicle's tail month", () => {
    overviewState.rows = [
      row({
        id: 1,
        status: "needs_generation",
        open_before: 6,
        open_before_month: "2026-07",
        wb_count: 0,
      }),
    ];
    renderOverview();
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2026-07-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2026-07-31" } });
    fireEvent.click(screen.getByRole("button", { name: "Машина 1" }));
    fireEvent.click(screen.getByTestId("journal-generate-btn"));

    const dialog = screen.getByRole("dialog");
    expect(within(dialog).getByText("Генерация: Машина 1")).toBeInTheDocument();
    expect(within(dialog).getByLabelText("Период с")).toHaveValue("2026-07-01");
    expect(within(dialog).getByLabelText("Период по")).toHaveValue("2026-07-31");
  });
});

describe("FleetOverviewView generate period", () => {
  afterEach(() => {
    cleanup();
  });

  it("pages the overview period from the journal and opens generate for that period", () => {
    overviewState.rows = [row({ id: 4, status: "has_red_days" })];
    renderOverview();

    fireEvent.click(screen.getByRole("button", { name: "Машина 4" }));
    fireEvent.click(screen.getByRole("button", { name: "Предыдущий месяц" }));

    const current = currentMonthBounds().from.slice(0, 7);
    const expected = monthBounds(shiftMonth(current, -1));
    expect(screen.getByLabelText("Период с")).toHaveValue(expected.from);
    expect(screen.getByLabelText("Период по")).toHaveValue(expected.to);

    fireEvent.click(screen.getByTestId("journal-generate-btn"));

    const dialog = screen.getByRole("dialog");
    expect(within(dialog).getByLabelText("Период с")).toHaveValue(expected.from);
    expect(within(dialog).getByLabelText("Период по")).toHaveValue(expected.to);
  });

  it("keeps the top period when generating from the row button", () => {
    overviewState.rows = [row({ id: 4, status: "needs_generation", wb_count: 0 })];
    renderOverview();

    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать" }));

    const top = currentMonthBounds();
    const dialog = screen.getByRole("dialog");
    expect(within(dialog).getByLabelText("Период с")).toHaveValue(top.from);
    expect(within(dialog).getByLabelText("Период по")).toHaveValue(top.to);
  });
});
