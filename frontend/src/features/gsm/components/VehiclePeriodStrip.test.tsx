import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { VehiclePeriodStrip } from "@/features/gsm/components/VehiclePeriodStrip";
import type { GsmDriver, GsmWaybill } from "@/features/gsm/types/gsm";

const DRIVER: GsmDriver = {
  id: 7,
  full_name: "Кулигин Никита",
  license_number: "44 21 846315",
  license_issued_at: null,
  personnel_number: "143",
  snils: null,
  is_active: true,
};

const WAYBILLS: GsmWaybill[] = [
  {
    id: 1,
    vehicle_id: 1,
    date: "2025-04-03",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10000,
    odometer_end: 10190,
    fuel_start: 20,
    fuel_issued: 40,
    fuel_end: 42.1,
    km: 190,
    route: [{ from: "Кострома", to: "Ярославль", km: 190, station_id: 12 }],
    warnings: ["weekend_anchor"],
  },
  {
    id: 2,
    vehicle_id: 1,
    date: "2025-04-04",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10190,
    odometer_end: 10380,
    fuel_start: 42.1,
    fuel_issued: 0,
    fuel_end: 24.2,
    km: 190,
    route: [{ from: "Кострома", to: "Иваново", km: 190 }],
    warnings: [],
  },
  {
    id: 3,
    vehicle_id: 1,
    date: "2025-04-07",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10380,
    odometer_end: 10570,
    fuel_start: 24.2,
    fuel_issued: 35,
    fuel_end: 40,
    km: 190,
    route: [{ from: "Кострома", to: "Владимир", km: 190 }],
    warnings: ["hook_above_threshold"],
  },
];

describe("VehiclePeriodStrip", () => {
  afterEach(() => {
    cleanup();
  });

  it("renders day grid with route/km/driver, fuel and odometer, anchors and clickable warnings", () => {
    render(
      <VehiclePeriodStrip
        vehicleName="Geely Monjaro"
        plateNumber="A123BC77"
        tankVolumeLiters={60}
        waybills={WAYBILLS}
        driversById={new Map([[7, DRIVER]])}
      />,
    );

    expect(screen.getByText(/Geely Monjaro/)).toBeInTheDocument();
    expect(screen.getByLabelText("Таймлайн остатка бака")).toBeInTheDocument();

    const anchorDay = screen.getByLabelText("День 2025-04-03");
    expect(anchorDay).toHaveAttribute("data-anchor", "true");
    expect(within(anchorDay).getByText("Якорь")).toBeInTheDocument();
    expect(within(anchorDay).getByText(/Кострома → Ярославль/)).toBeInTheDocument();
    expect(within(anchorDay).getByText(/190 км/)).toBeInTheDocument();
    expect(within(anchorDay).getByText(/Кулигин Никита/)).toBeInTheDocument();
    expect(within(anchorDay).getByText(/Топливо:/)).toBeInTheDocument();
    expect(within(anchorDay).getByText(/Одометр:/)).toBeInTheDocument();

    const burnDay = screen.getByLabelText("День 2025-04-04");
    expect(burnDay).toHaveAttribute("data-anchor", "false");

    const weekendBadge = within(anchorDay).getByRole("button", { name: /Предупреждение: Выходной/ });
    fireEvent.click(weekendBadge);
    expect(within(anchorDay).getByRole("status")).toHaveTextContent(/выходной или праздник/i);

    const hookDay = screen.getByLabelText("День 2025-04-07");
    const hookBadge = within(hookDay).getByRole("button", { name: /Предупреждение: Крюк/ });
    fireEvent.click(hookBadge);
    expect(within(hookDay).getByRole("status")).toHaveTextContent(/порог/i);
  });

  it("marks manual_intervention days as problematic, clickable, and shows balance_route badge", () => {
    const onDayClick = vi.fn();
    const waybills: GsmWaybill[] = [
      {
        ...WAYBILLS[1],
        id: 10,
        date: "2025-04-05",
        fuel_issued: 50,
        warnings: ["manual_intervention"],
      },
      {
        ...WAYBILLS[1],
        id: 11,
        date: "2025-04-06",
        fuel_issued: 30,
        warnings: ["balance_route"],
      },
    ];

    render(
      <VehiclePeriodStrip
        vehicleName="Geely Monjaro"
        plateNumber="A123BC77"
        tankVolumeLiters={60}
        waybills={waybills}
        driversById={new Map([[7, DRIVER]])}
        onDayClick={onDayClick}
      />,
    );

    const problemDay = screen.getByLabelText("День 2025-04-05");
    expect(problemDay).toHaveAttribute("data-problematic", "true");
    expect(problemDay).toHaveStyle({ border: "1px solid #f04438", background: "#fef3f2" });
    expect(problemDay).toHaveAttribute("role", "button");

    const manualBadge = within(problemDay).getByRole("button", {
      name: /Предупреждение: Ручная доработка/,
    });
    fireEvent.click(manualBadge);
    expect(within(problemDay).getByRole("status")).toHaveTextContent(/баланс|ручн/i);

    fireEvent.click(problemDay);
    expect(onDayClick).toHaveBeenCalledWith(expect.objectContaining({ date: "2025-04-05", id: 10 }));

    const balanceDay = screen.getByLabelText("День 2025-04-06");
    expect(balanceDay).toHaveAttribute("data-problematic", "false");
    const balanceBadge = within(balanceDay).getByRole("button", {
      name: /Предупреждение: Маршрут для баланса/,
    });
    fireEvent.click(balanceBadge);
    expect(within(balanceDay).getByRole("status")).toHaveTextContent(/удлин|маршрут/i);
  });
});
