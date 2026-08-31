import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { VehicleWaybillJournal } from "@/features/gsm/components/VehicleWaybillJournal";
import { VehicleGenerateDialog } from "@/features/gsm/components/VehicleGenerateDialog";
import { ApiError } from "@/shared/lib/apiError";
import type { FleetOverviewRow, GsmTransaction, GsmWaybill } from "@/features/gsm/types/gsm";

const mockGenerate = vi.fn();

vi.mock("@/features/gsm/components/ManualWaybillDialog", () => ({
  ManualWaybillDialog: () => null,
}));
vi.mock("@/features/gsm/components/WaybillDayDrawer", () => ({
  WaybillDayDrawer: ({
    open,
    waybill,
  }: {
    open: boolean;
    waybill: { date?: string } | null;
  }) => (open ? <div data-testid="waybill-drawer">{waybill?.date}</div> : null),
}));
vi.mock("@/features/gsm/components/VehiclePeriodStrip", () => ({
  VehiclePeriodStrip: () => <div>лента</div>,
}));

const WAYBILLS: GsmWaybill[] = [
  {
    id: 1,
    vehicle_id: 4,
    date: "2026-08-03",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10000,
    odometer_end: 10200,
    fuel_start: 20,
    fuel_issued: 40,
    fuel_end: 24,
    km: 200,
    route: [{ from: "Завод", to: "Объект", km: 200 }],
    warnings: ["manual_intervention"],
    warning_details: [{ code: "manual_intervention", detail: "бак не сходится" }],
  },
  {
    id: 2,
    vehicle_id: 4,
    date: "2026-08-04",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10200,
    odometer_end: 10280,
    fuel_start: 24,
    fuel_issued: 10,
    fuel_end: 20,
    km: 80,
    route: [{ from: "Завод", to: "Склад", km: 80 }],
    warnings: [],
  },
];

const TX_AUG: GsmTransaction[] = [
  {
    ts: "2026-08-05T10:00:00",
    card_number: "1",
    vehicle_id: 4,
    service_type: "fuel",
    fuel_grade: "ДТ",
    qty_liters: 40,
    amount: 1000,
    station_id: null,
    address: null,
  },
];

const journalState = {
  waybills: WAYBILLS as GsmWaybill[],
  transactions: TX_AUG as GsmTransaction[],
  txError: null as Error | null,
  waybillParams: [] as unknown[],
  txParams: [] as unknown[],
};

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmWaybillsQuery: (params: unknown) => {
    journalState.waybillParams.push(params);
    return {
      isLoading: false,
      error: null,
      data: journalState.waybills,
    };
  },
  useGsmTransactionsQuery: (params: unknown) => {
    journalState.txParams.push(params);
    return {
      isLoading: false,
      error: journalState.txError,
      data: journalState.txError
        ? undefined
        : { rows: journalState.transactions, total_count: journalState.transactions.length, sum_liters: 0, sum_amount: 0 },
    };
  },
  useGsmDriversQuery: () => ({
    isLoading: false,
    error: null,
    data: [{ id: 7, full_name: "Кулигин", license_number: "1", is_active: true }],
  }),
  useGsmVehiclesQuery: () => ({
    isLoading: false,
    error: null,
    data: [
      {
        id: 4,
        name: "Tugella",
        plate_number: "О 848",
        tank_volume_liters: 55,
        norm_summer: 9.4,
        norm_winter: 10.3,
        primary_driver_id: 7,
        is_active: true,
      },
    ],
  }),
  useGenerateGsmWaybillsMutation: () => ({
    mutateAsync: mockGenerate,
    isPending: false,
  }),
}));

const ROW: FleetOverviewRow = {
  vehicle: { id: 4, name: "Tugella", plate_number: "О 848" },
  tx_count: 1,
  tx_liters: 50,
  tx_amount: 1,
  tx_last_date: "2026-08-04",
  wb_count: 2,
  wb_km: 280,
  wb_fuel_issued: 50,
  wb_last_date: "2026-08-04",
  red_days: 1,
  draft_count: 2,
  confirmed_count: 0,
  exported_count: 0,
  fuel_end_last: 20,
  liters_diff: 0,
  open_before: 0,
  open_before_month: null,
  chain_broken: false,
  status: "has_red_days",
};

const renderJournal = (
  props: Partial<Parameters<typeof VehicleWaybillJournal>[0]> = {},
) =>
  render(
    <VehicleWaybillJournal
      vehicleId={4}
      vehicleName="Tugella"
      plateNumber="О 848"
      periodFrom="2026-08-01"
      periodTo="2026-08-31"
      {...props}
    />,
  );

describe("VehicleWaybillJournal", () => {
  beforeEach(() => {
    journalState.waybills = WAYBILLS;
    journalState.transactions = TX_AUG;
    journalState.txError = null;
    journalState.waybillParams = [];
    journalState.txParams = [];
  });

  afterEach(() => {
    cleanup();
  });

  it("renders the feed with a summary line instead of the waybill table", () => {
    renderJournal();
    expect(screen.getByTestId("feed-summary")).toHaveTextContent(
      "Итого за август 2026 г.: 2 ПЛ, 280 км, выдано 50 л",
    );
    expect(screen.getByTestId("feed-day-2026-08-03")).toBeInTheDocument();
    expect(screen.getByTestId("feed-day-2026-08-05")).toBeInTheDocument();
    expect(screen.queryByRole("table")).not.toBeInTheDocument();
  });

  it("shows the warning detail badge on the red waybill card", () => {
    renderJournal();
    expect(screen.getByTestId("feed-wb-1")).toHaveTextContent(/бак не сходится/);
  });

  it("marks tx-only day as gap and focuses Generate without calling onGenerate", () => {
    const onGenerate = vi.fn();
    const scrollIntoView = vi.fn();
    HTMLElement.prototype.scrollIntoView = scrollIntoView;

    renderJournal({ onGenerate });

    const gap = screen.getByTestId("cal-day-2026-08-05");
    expect(gap).toHaveAttribute("aria-label", expect.stringMatching(/нет путевого/i));

    fireEvent.click(gap);
    expect(onGenerate).not.toHaveBeenCalled();
    expect(screen.getByTestId("journal-generate-btn")).toHaveFocus();
    expect(scrollIntoView).toHaveBeenCalled();
  });

  it("gap CTA in the feed focuses Generate without calling onGenerate", () => {
    const onGenerate = vi.fn();
    HTMLElement.prototype.scrollIntoView = vi.fn();

    renderJournal({ onGenerate });

    const gapCard = screen.getByTestId("feed-gap-2026-08-05");
    fireEvent.click(within(gapCard).getByRole("button", { name: "Сгенерировать" }));
    expect(onGenerate).not.toHaveBeenCalled();
    expect(screen.getByTestId("journal-generate-btn")).toHaveFocus();
  });

  it("opens drawer path when clicking a day with PL", () => {
    renderJournal();
    fireEvent.click(screen.getByTestId("cal-day-2026-08-03"));
    expect(screen.getByTestId("waybill-drawer")).toHaveTextContent("2026-08-03");
  });

  it("opens the drawer when clicking a waybill card in the feed", () => {
    renderJournal();
    fireEvent.click(screen.getByTestId("feed-wb-2"));
    expect(screen.getByTestId("waybill-drawer")).toHaveTextContent("2026-08-04");
  });

  it("shows the calendar month from periodFrom, not a local month", () => {
    renderJournal();
    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Август 2026 г.");
    expect(journalState.waybillParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-08-01",
      periodTo: "2026-08-31",
    });
  });

  it("pages to the previous month and refetches after the parent applies bounds", () => {
    const onPeriodChange = vi.fn();
    const { rerender } = renderJournal({ onPeriodChange });
    expect(journalState.waybillParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-08-01",
      periodTo: "2026-08-31",
    });

    fireEvent.click(screen.getByRole("button", { name: "Предыдущий месяц" }));

    expect(onPeriodChange).toHaveBeenCalledWith({ from: "2026-07-01", to: "2026-07-31" });
    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Август 2026 г.");
    expect(journalState.waybillParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-08-01",
      periodTo: "2026-08-31",
    });

    rerender(
      <VehicleWaybillJournal
        vehicleId={4}
        vehicleName="Tugella"
        plateNumber="О 848"
        periodFrom="2026-07-01"
        periodTo="2026-07-31"
        onPeriodChange={onPeriodChange}
      />,
    );

    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Июль 2026 г.");
    expect(journalState.waybillParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-07-01",
      periodTo: "2026-07-31",
    });
    expect(journalState.txParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-07-01",
      periodTo: "2026-07-31",
    });
  });

  it("sets the visible month from periodFrom and queries the exact period props", () => {
    const { rerender } = renderJournal();
    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Август 2026 г.");
    expect(journalState.waybillParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-08-01",
      periodTo: "2026-08-31",
    });

    rerender(
      <VehicleWaybillJournal
        vehicleId={4}
        vehicleName="Tugella"
        plateNumber="О 848"
        periodFrom="2026-07-15"
        periodTo="2026-07-20"
      />,
    );
    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Июль 2026 г.");
    expect(journalState.waybillParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-07-15",
      periodTo: "2026-07-20",
    });
    expect(journalState.txParams.at(-1)).toEqual({
      vehicleId: 4,
      periodFrom: "2026-07-15",
      periodTo: "2026-07-20",
    });
  });

  it("notifies parent on paging and generates with current period props", () => {
    const onGenerate = vi.fn();
    const onPeriodChange = vi.fn();
    const { rerender } = renderJournal({ onGenerate, onPeriodChange });

    fireEvent.click(screen.getByRole("button", { name: "Предыдущий месяц" }));
    expect(onPeriodChange).toHaveBeenCalledWith({ from: "2026-07-01", to: "2026-07-31" });

    rerender(
      <VehicleWaybillJournal
        vehicleId={4}
        vehicleName="Tugella"
        plateNumber="О 848"
        periodFrom="2026-07-01"
        periodTo="2026-07-31"
        onGenerate={onGenerate}
        onPeriodChange={onPeriodChange}
      />,
    );
    fireEvent.click(screen.getByTestId("journal-generate-btn"));

    expect(onGenerate).toHaveBeenCalledWith({ from: "2026-07-01", to: "2026-07-31" });
  });

  it("calls onGenerate with the exact period props, not full month bounds", () => {
    const onGenerate = vi.fn();
    renderJournal({
      periodFrom: "2026-08-10",
      periodTo: "2026-08-20",
      onGenerate,
    });

    fireEvent.click(screen.getByTestId("journal-generate-btn"));

    expect(onGenerate).toHaveBeenCalledWith({ from: "2026-08-10", to: "2026-08-20" });
  });

  it("shows empty calendar message for July without tx or PL; feed stays empty", () => {
    journalState.waybills = [];
    journalState.transactions = [];
    render(
      <VehicleWaybillJournal
        vehicleId={4}
        vehicleName="Palisade"
        plateNumber="А 001"
        periodFrom="2026-07-01"
        periodTo="2026-07-31"
      />,
    );
    expect(screen.getByText(/нет движений/i)).toBeInTheDocument();
    expect(screen.getByTestId("cal-day-2026-07-01")).toBeInTheDocument();
    expect(screen.getByTestId("cal-day-2026-07-31")).toBeInTheDocument();
    expect(screen.getByTestId("journal-4")).toBeInTheDocument();
    expect(screen.queryByTestId("vehicle-day-feed")).not.toBeInTheDocument();
  });
});

describe("VehicleGenerateDialog", () => {
  afterEach(() => {
    cleanup();
    mockGenerate.mockReset();
  });

  it("generates without override when history exists", async () => {
    mockGenerate.mockResolvedValue({
      waybills: [],
      warnings: [],
      days_created: 3,
      problematic_days: [],
      manual_days: 0,
    });
    render(
      <VehicleGenerateDialog
        open
        row={ROW}
        periodFrom="2026-08-01"
        periodTo="2026-08-31"
        onClose={() => undefined}
      />,
    );
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать" }));
    await waitFor(() => expect(mockGenerate).toHaveBeenCalled());
    expect(mockGenerate.mock.calls[0][0]).toMatchObject({
      vehicle_id: 4,
      period_from: "2026-08-01",
      period_to: "2026-08-31",
      fuel_start: null,
      odometer_start: null,
    });
    expect(screen.getByText(/Создано 3 дней/)).toBeInTheDocument();
  });

  it("highlights start fields on gsm_start_required", async () => {
    mockGenerate.mockRejectedValue(
      new ApiError("нужен старт", 422, "нужен старт", "gsm_start_required"),
    );
    render(
      <VehicleGenerateDialog
        open
        row={ROW}
        periodFrom="2026-08-01"
        periodTo="2026-08-31"
        onClose={() => undefined}
      />,
    );
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать" }));
    await waitFor(() => expect(screen.getByText("нужен старт")).toBeInTheDocument());
    expect(screen.getByLabelText("Стартовый остаток бака").parentElement).toHaveStyle(
      "outline: 2px solid #f04438",
    );
  });
});
