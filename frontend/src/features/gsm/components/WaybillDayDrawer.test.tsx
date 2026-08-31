import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WaybillDayDrawer } from "@/features/gsm/components/WaybillDayDrawer";
import type { GsmDriver, GsmRoute, GsmStation, GsmVehicle, GsmWaybill } from "@/features/gsm/types/gsm";

const mockPatch = vi.fn();

const VEHICLE: GsmVehicle = {
  id: 1,
  name: "Geely Monjaro",
  plate_number: "A123BC77",
  tank_volume_liters: 60,
  norm_summer: 10,
  norm_winter: 12,
  primary_driver_id: 7,
  is_active: true,
};

const DRIVER: GsmDriver = {
  id: 7,
  full_name: "Кулигин Никита",
  license_number: "44 21",
  license_issued_at: null,
  personnel_number: null,
  snils: null,
  is_active: true,
};

const STATION: GsmStation = {
  id: 12,
  address: "Кострома, АЗС",
  brand: "ТНК",
  lat: null,
  lon: null,
  geocode_source: null,
};

const ROUTES: GsmRoute[] = [
  {
    id: 5,
    vehicle_id: 1,
    addr_a: "Кострома",
    addr_b: "Ярославль",
    km: 150,
    frequency: 3,
    typical_station_ids: [12],
  },
  {
    id: 6,
    vehicle_id: 1,
    addr_a: "Кострома",
    addr_b: "Иваново",
    km: 190,
    frequency: 1,
    typical_station_ids: [],
  },
  {
    id: 59,
    vehicle_id: 1,
    addr_a: "г.Владимир, ул.Добросельская",
    addr_b: "Кострома, ул. Кузнецкая, д.18Б",
    km: 225,
    frequency: 8,
    typical_station_ids: [],
  },
  {
    id: 64,
    vehicle_id: 1,
    addr_a: "Кострома, ул. Кузнецкая, д.18Б",
    addr_b: "г.Владимир, ул.Добросельская",
    km: 225,
    frequency: 2,
    typical_station_ids: [],
  },
];

const WAYBILLS: GsmWaybill[] = [
  {
    id: 11,
    vehicle_id: 1,
    date: "2025-04-03",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10000,
    odometer_end: 10200,
    fuel_start: 20,
    fuel_issued: 40,
    fuel_end: 40,
    km: 200,
    route: [{ from: "A", to: "B", km: 200 }],
    warnings: [],
  },
  {
    id: 12,
    vehicle_id: 1,
    date: "2025-04-04",
    driver_id: 7,
    status: "draft",
    source: "auto",
    odometer_start: 10200,
    odometer_end: 10400,
    fuel_start: 40,
    fuel_issued: 0,
    fuel_end: 20,
    km: 200,
    route: [{ from: "A", to: "C", km: 200 }],
    warnings: [],
  },
];

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmDriversQuery: () => ({ isLoading: false, error: null, data: [DRIVER] }),
  useGsmRoutesQuery: () => ({ isLoading: false, error: null, data: ROUTES }),
  useGsmStationsQuery: () => ({ isLoading: false, error: null, data: [STATION] }),
  useGsmSettingsQuery: () => ({
    isLoading: false,
    error: null,
    data: {
      winter_start: "11-01",
      hook_threshold_km: 13,
      season_mode: "summer",
      season_switched_at: null,
    },
  }),
  usePatchGsmWaybillMutation: () => ({
    mutateAsync: mockPatch,
    isPending: false,
    reset: vi.fn(),
  }),
}));

describe("WaybillDayDrawer", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("filters routes by АЗС, shows downstream preview, and PATCHes on save", async () => {
    const onSaved = vi.fn();
    const onClose = vi.fn();
    mockPatch.mockResolvedValue({ ...WAYBILLS[0], km: 150 });

    render(
      <WaybillDayDrawer
        open
        waybill={WAYBILLS[0]}
        vehicle={VEHICLE}
        periodWaybills={WAYBILLS}
        onClose={onClose}
        onSaved={onSaved}
      />,
    );

    expect(screen.getByText(/Правка дня 2025-04-03/)).toBeInTheDocument();

    fireEvent.change(screen.getByLabelText("Фильтр по АЗС"), { target: { value: "12" } });
    const routeSelect = screen.getByLabelText("Маршрут из библиотеки");
    expect(routeSelect).toHaveTextContent(/Ярославль/);
    expect(routeSelect).not.toHaveTextContent(/Иваново/);

    fireEvent.change(routeSelect, { target: { value: "5" } });
    expect(screen.getByLabelText("Км за день")).toHaveValue(150);

    expect(await screen.findByLabelText("Превью пересчёта")).toBeInTheDocument();
    expect(screen.getByText("2025-04-04")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Сохранить" }));

    await waitFor(() => {
      expect(mockPatch).toHaveBeenCalledWith({
        id: 11,
        payload: {
          driver_id: 7,
          km: 150,
          route: [
            {
              from: "Кострома",
              to: "Ярославль",
              km: 150,
              route_id: 5,
              station_id: 12,
            },
          ],
        },
      });
    });
    expect(onSaved).toHaveBeenCalled();
    expect(onClose).toHaveBeenCalled();
  });

  it("orients remote-first library row so PATCH starts at home with twin route_id", async () => {
    const onSaved = vi.fn();
    const onClose = vi.fn();
    mockPatch.mockResolvedValue({ ...WAYBILLS[0], km: 225 });

    render(
      <WaybillDayDrawer
        open
        waybill={WAYBILLS[0]}
        vehicle={VEHICLE}
        periodWaybills={WAYBILLS}
        onClose={onClose}
        onSaved={onSaved}
      />,
    );

    fireEvent.change(screen.getByLabelText("Маршрут из библиотеки"), { target: { value: "59" } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить" }));

    await waitFor(() => {
      expect(mockPatch).toHaveBeenCalledWith({
        id: 11,
        payload: {
          driver_id: 7,
          km: 225,
          route: [
            {
              from: "Кострома, ул. Кузнецкая, д.18Б",
              to: "г.Владимир, ул.Добросельская",
              km: 225,
              route_id: 64,
              station_id: null,
            },
          ],
        },
      });
    });
    expect(onSaved).toHaveBeenCalled();
    expect(onClose).toHaveBeenCalled();
  });
});
