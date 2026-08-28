import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { VehicleDayFeed } from "@/features/gsm/components/VehicleDayFeed";
import { buildVehicleDayFeed } from "@/features/gsm/lib/vehicleDayFeed";
import type { GsmDriver, GsmTransaction, GsmWaybill } from "@/features/gsm/types/gsm";

const wb = (overrides: Partial<GsmWaybill> = {}): GsmWaybill => ({
  id: 1,
  vehicle_id: 1,
  date: "2026-08-05",
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
  warnings: [],
  ...overrides,
});

const tx = (overrides: Partial<GsmTransaction> = {}): GsmTransaction => ({
  ts: "2026-08-05T10:30:00",
  card_number: "1",
  vehicle_id: 1,
  service_type: "fuel",
  fuel_grade: "ДТ",
  qty_liters: 40,
  amount: 1234.5,
  station_id: null,
  address: "АЗС Лукойл 12",
  ...overrides,
});

const DRIVERS: GsmDriver[] = [
  { id: 7, full_name: "Кулигин", license_number: "1", is_active: true } as GsmDriver,
];
const driversById = new Map(DRIVERS.map((d) => [d.id, d]));

const renderFeed = (
  waybills: GsmWaybill[],
  transactions: GsmTransaction[],
  props: { onGapClick?: () => void; onWaybillClick?: (wb: GsmWaybill) => void } = {},
) => {
  const feed = buildVehicleDayFeed("2026-08-01", "2026-08-31", waybills, transactions);
  return render(
    <VehicleDayFeed month="2026-08" feed={feed} driversById={driversById} {...props} />,
  );
};

describe("VehicleDayFeed", () => {
  afterEach(() => {
    cleanup();
  });

  it("renders a tx card with time, service, station, liters and amount", () => {
    renderFeed([wb({ id: 2 })], [tx()]);
    const section = screen.getByTestId("feed-day-2026-08-05");
    expect(section).toHaveTextContent("10:30");
    expect(section).toHaveTextContent("Топливо");
    expect(section).toHaveTextContent("АЗС Лукойл 12");
    expect(section).toHaveTextContent(/40\s?л/);
    expect(section).toHaveTextContent(/1\s?234,50\s?₽/);
  });

  it("renders an 'other' tx as a regular card, not a gap", () => {
    renderFeed([], [tx({ service_type: "other", address: null })]);
    const section = screen.getByTestId("feed-day-2026-08-05");
    expect(section).toHaveTextContent("other");
    expect(section).toHaveTextContent("—");
    expect(screen.queryByRole("button", { name: /Сгенерировать/i })).not.toBeInTheDocument();
    expect(
      screen.queryByRole("region", { name: /нет путевого на заправку\/мойку/i }),
    ).not.toBeInTheDocument();
  });

  it("renders an amber gap card with CTA that calls onGapClick without generating", () => {
    const onGapClick = vi.fn();
    renderFeed([], [tx({ service_type: "wash", qty_liters: null, amount: 500 })], {
      onGapClick,
    });
    const gap = screen.getByTestId("feed-gap-2026-08-05");
    expect(gap).toHaveStyle({ background: "#fffaeb" });
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать" }));
    expect(onGapClick).toHaveBeenCalledTimes(1);
  });

  it("hides the gap CTA when onGapClick is not provided", () => {
    renderFeed([], [tx()]);
    expect(screen.getByTestId("feed-gap-2026-08-05")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Сгенерировать" })).not.toBeInTheDocument();
  });

  it("renders a waybill card and forwards clicks to onWaybillClick", () => {
    const onWaybillClick = vi.fn();
    const waybill = wb({ id: 3 });
    renderFeed([waybill], [], { onWaybillClick });
    const card = screen.getByTestId("feed-wb-3");
    expect(card).toHaveTextContent("Кулигин");
    expect(card).toHaveTextContent("Завод → Объект");
    expect(card).toHaveTextContent(/200\s?км/);
    expect(card).toHaveTextContent("draft");
    fireEvent.click(card);
    expect(onWaybillClick).toHaveBeenCalledWith(waybill);
  });

  it("shows a red badge for manual_intervention without turning the day into a gap", () => {
    renderFeed(
      [
        wb({
          id: 4,
          warnings: ["manual_intervention"],
          warning_details: [{ code: "manual_intervention", detail: "бак не сходится" }],
        }),
      ],
      [tx()],
    );
    expect(screen.getByTestId("feed-wb-4")).toHaveTextContent(/Ручная доработка/);
    expect(screen.queryByTestId("feed-gap-2026-08-05")).not.toBeInTheDocument();
  });

  it("renders the summary line with waybill totals of the month", () => {
    renderFeed(
      [wb({ id: 1, km: 200, fuel_issued: 40 }), wb({ id: 2, date: "2026-08-06", km: 80, fuel_issued: 10 })],
      [],
    );
    expect(screen.getByTestId("feed-summary")).toHaveTextContent(
      "Итого за август 2026 г.: 2 ПЛ, 280 км, выдано 50 л",
    );
  });

  it("renders nothing for an empty feed", () => {
    const { container } = renderFeed([], []);
    expect(container).toBeEmptyDOMElement();
  });
});
