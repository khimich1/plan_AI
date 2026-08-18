import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ManualWaybillDialog } from "@/features/gsm/components/ManualWaybillDialog";
import type { GsmDriver, GsmRoute, GsmStation, GsmVehicle, GsmWaybill } from "@/features/gsm/types/gsm";

const mockCreate = vi.fn();

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
];

const PREV: GsmWaybill = {
  id: 10,
  vehicle_id: 1,
  date: "2025-04-02",
  driver_id: 7,
  status: "draft",
  source: "auto",
  odometer_start: 9900,
  odometer_end: 10000,
  fuel_start: 10,
  fuel_issued: 0,
  fuel_end: 25.5,
  km: 100,
  route: [{ from: "A", to: "B", km: 100 }],
  warnings: [],
};

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmVehiclesQuery: () => ({ isLoading: false, error: null, data: [VEHICLE] }),
  useGsmDriversQuery: () => ({ isLoading: false, error: null, data: [DRIVER] }),
  useGsmStationsQuery: () => ({ isLoading: false, error: null, data: [STATION] }),
  useGsmRoutesQuery: () => ({ isLoading: false, error: null, data: ROUTES }),
  useCreateGsmWaybillMutation: () => ({
    mutateAsync: mockCreate,
    isPending: false,
    reset: vi.fn(),
  }),
}));

describe("ManualWaybillDialog", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("autofills fuel/odo from previous day and POSTs create payload", async () => {
    const onCreated = vi.fn();
    const onClose = vi.fn();
    mockCreate.mockResolvedValue({ id: 99, ...PREV, date: "2025-04-03" });

    render(
      <ManualWaybillDialog
        open
        onClose={onClose}
        defaultVehicleId={1}
        defaultDate="2025-04-03"
        periodWaybills={[PREV]}
        onCreated={onCreated}
      />,
    );

    expect(await screen.findByLabelText("Остаток бака")).toHaveValue("25.5");
    expect(screen.getByLabelText("Одометр старт")).toHaveValue("10000");
    expect(screen.getByLabelText("Водитель")).toHaveValue("7");

    fireEvent.change(screen.getByLabelText("Маршрут"), { target: { value: "5" } });
    fireEvent.change(screen.getByLabelText("Выдано"), { target: { value: "30" } });
    fireEvent.click(screen.getByRole("button", { name: "Создать ПЛ" }));

    await waitFor(() => {
      expect(mockCreate).toHaveBeenCalledWith({
        vehicle_id: 1,
        date: "2025-04-03",
        driver_id: 7,
        route: [
          {
            from: "Кострома",
            to: "Ярославль",
            km: 150,
            route_id: 5,
            station_id: 12,
          },
        ],
        fuel_issued: 30,
        fuel_start: 25.5,
        odometer_start: 10000,
      });
    });
    expect(onCreated).toHaveBeenCalled();
    expect(onClose).toHaveBeenCalled();
  });
});
