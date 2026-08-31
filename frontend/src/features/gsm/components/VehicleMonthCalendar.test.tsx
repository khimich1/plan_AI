import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { VehicleMonthCalendar } from "@/features/gsm/components/VehicleMonthCalendar";
import type { VehicleDayCell } from "@/features/gsm/lib/vehicleDayMap";
import type { GsmWaybill } from "@/features/gsm/types/gsm";

const wb = (overrides: Partial<GsmWaybill> = {}): GsmWaybill => ({
  id: 1,
  vehicle_id: 1,
  date: "2026-08-03",
  driver_id: 1,
  status: "draft",
  source: "auto",
  odometer_start: null,
  odometer_end: null,
  fuel_start: null,
  fuel_issued: null,
  fuel_end: null,
  km: 0,
  route: [],
  warnings: [],
  ...overrides,
});

const cell = (overrides: Partial<VehicleDayCell>): VehicleDayCell => ({
  date: "2026-08-01",
  hasTx: false,
  hasPl: false,
  isGap: false,
  isRed: false,
  waybill: null,
  ...overrides,
});

describe("VehicleMonthCalendar", () => {
  afterEach(() => {
    cleanup();
  });

  it("renders day cells with test ids and weekday headers", () => {
    render(
      <VehicleMonthCalendar
        cells={[
          cell({ date: "2026-08-01" }),
          cell({ date: "2026-08-02" }),
        ]}
      />,
    );
    expect(screen.getByTestId("cal-day-2026-08-01")).toBeInTheDocument();
    expect(screen.getByTestId("cal-day-2026-08-02")).toBeInTheDocument();
    expect(screen.getByText("Пн")).toBeInTheDocument();
    expect(screen.getByText("Вс")).toBeInTheDocument();
  });

  it("distinguishes gap from red PL by accessible name", () => {
    const gapWb = null;
    const redWb = wb({
      id: 2,
      date: "2026-08-03",
      warnings: ["manual_intervention"],
    });
    render(
      <VehicleMonthCalendar
        cells={[
          cell({
            date: "2026-08-05",
            hasTx: true,
            isGap: true,
            waybill: gapWb,
          }),
          cell({
            date: "2026-08-03",
            hasTx: true,
            hasPl: true,
            isRed: true,
            waybill: redWb,
          }),
        ]}
      />,
    );
    expect(
      screen.getByRole("button", { name: /нет путевого на заправку\/мойку/i }),
    ).toBeInTheDocument();
    const red = screen.getByTestId("cal-day-2026-08-03");
    expect(red).toHaveAttribute("aria-label", expect.stringMatching(/ручн|красн|доработ/i));
    expect(red.getAttribute("aria-label")).not.toMatch(/нет путевого/i);
  });

  it("calls onDayClick for PL day and onGapClick for gap", () => {
    const onDayClick = vi.fn();
    const onGapClick = vi.fn();
    const waybill = wb({ id: 3, date: "2026-08-04" });
    render(
      <VehicleMonthCalendar
        cells={[
          cell({ date: "2026-08-04", hasPl: true, waybill }),
          cell({ date: "2026-08-05", hasTx: true, isGap: true }),
        ]}
        onDayClick={onDayClick}
        onGapClick={onGapClick}
      />,
    );
    fireEvent.click(screen.getByTestId("cal-day-2026-08-04"));
    expect(onDayClick).toHaveBeenCalledWith(waybill);
    expect(onGapClick).not.toHaveBeenCalled();

    fireEvent.click(screen.getByTestId("cal-day-2026-08-05"));
    expect(onGapClick).toHaveBeenCalledTimes(1);
    expect(onDayClick).toHaveBeenCalledTimes(1);
  });

  it("does not make empty days or padding slots clickable buttons", () => {
    render(
      <VehicleMonthCalendar
        cells={[cell({ date: "2026-08-01" })]}
        onDayClick={vi.fn()}
        onGapClick={vi.fn()}
      />,
    );
    const empty = screen.getByTestId("cal-day-2026-08-01");
    expect(empty.tagName.toLowerCase()).not.toBe("button");
    expect(screen.queryAllByRole("button")).toHaveLength(0);
  });

  it("shows empty-period text while keeping range cells", () => {
    render(
      <VehicleMonthCalendar
        cells={[
          cell({ date: "2026-07-01" }),
          cell({ date: "2026-07-02" }),
        ]}
      />,
    );
    expect(screen.getByText(/нет движений/i)).toBeInTheDocument();
    expect(screen.getByTestId("cal-day-2026-07-01")).toBeInTheDocument();
    expect(screen.getByTestId("cal-day-2026-07-02")).toBeInTheDocument();
  });

  it("hides empty-period text when any day has tx or PL", () => {
    render(
      <VehicleMonthCalendar
        cells={[
          cell({ date: "2026-08-01" }),
          cell({ date: "2026-08-02", hasPl: true, waybill: wb() }),
        ]}
      />,
    );
    expect(screen.queryByText(/нет движений/i)).not.toBeInTheDocument();
  });
});

describe("VehicleMonthCalendar month header", () => {
  afterEach(() => {
    cleanup();
  });

  it("renders the month label capitalized and calls onMonthChange with ±1", () => {
    const onMonthChange = vi.fn();
    render(
      <VehicleMonthCalendar
        cells={[cell({ date: "2026-08-01" })]}
        month="2026-08"
        onMonthChange={onMonthChange}
      />,
    );
    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Август 2026 г.");

    fireEvent.click(screen.getByRole("button", { name: "Предыдущий месяц" }));
    expect(onMonthChange).toHaveBeenCalledWith(-1);
    fireEvent.click(screen.getByRole("button", { name: "Следующий месяц" }));
    expect(onMonthChange).toHaveBeenCalledWith(1);
  });

  it("shows January label across the year boundary", () => {
    render(
      <VehicleMonthCalendar
        cells={[cell({ date: "2026-01-15" })]}
        month="2026-01"
        onMonthChange={() => undefined}
      />,
    );
    expect(screen.getByTestId("cal-month-label")).toHaveTextContent("Январь 2026 г.");
  });

  it("keeps grid and day clicks intact with the header present", () => {
    const onDayClick = vi.fn();
    const waybill = wb({ id: 3, date: "2026-08-04" });
    render(
      <VehicleMonthCalendar
        cells={[cell({ date: "2026-08-04", hasPl: true, waybill })]}
        month="2026-08"
        onMonthChange={() => undefined}
        onDayClick={onDayClick}
      />,
    );
    fireEvent.click(screen.getByTestId("cal-day-2026-08-04"));
    expect(onDayClick).toHaveBeenCalledWith(waybill);
    expect(screen.getByText("Пн")).toBeInTheDocument();
  });

  it("does not render the header without month/onMonthChange", () => {
    render(<VehicleMonthCalendar cells={[cell({ date: "2026-08-01" })]} />);
    expect(screen.queryByTestId("cal-month-label")).not.toBeInTheDocument();
    expect(screen.queryAllByRole("button")).toHaveLength(0);
  });
});
