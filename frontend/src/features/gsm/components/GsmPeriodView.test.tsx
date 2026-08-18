import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { GsmPeriodView } from "@/features/gsm/components/GsmPeriodView";
import { ApiError } from "@/shared/lib/apiError";
import type { GsmDriver, GsmVehicle, GsmWaybill, WaybillGenerateResult } from "@/features/gsm/types/gsm";

const mockGenerate = vi.fn();
const mockExport = vi.fn();
const mockRefetch = vi.fn();
const mockUseWaybills = vi.fn();

const VEHICLE: GsmVehicle = {
  id: 1,
  name: "Geely Monjaro",
  plate_number: "A123BC77",
  tank_volume_liters: 60,
  norm_summer: 12,
  norm_winter: 14,
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

const WAYBILL: GsmWaybill = {
  id: 11,
  vehicle_id: 1,
  date: "2025-04-03",
  driver_id: 7,
  status: "draft",
  source: "auto",
  odometer_start: 10000,
  odometer_end: 10190,
  fuel_start: 20,
  fuel_issued: 40,
  fuel_end: 42,
  km: 190,
  route: [{ from: "A", to: "B", km: 190 }],
  warnings: ["weekend_anchor"],
};

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmVehiclesQuery: () => ({
    isLoading: false,
    error: null,
    data: [VEHICLE],
  }),
  useGsmDriversQuery: () => ({
    isLoading: false,
    error: null,
    data: [DRIVER],
  }),
  useGsmWaybillsQuery: (...args: unknown[]) => mockUseWaybills(...args),
  useGenerateGsmWaybillsMutation: () => ({
    mutateAsync: mockGenerate,
    isPending: false,
  }),
  useExportGsmWaybillsMutation: () => ({
    mutateAsync: mockExport,
    isPending: false,
  }),
  useGsmRoutesQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGsmStationsQuery: () => ({ isLoading: false, error: null, data: [] }),
  useGsmSettingsQuery: () => ({
    isLoading: false,
    error: null,
    data: { winter_start: "11-01", hook_threshold_km: 13 },
  }),
  usePatchGsmWaybillMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    reset: vi.fn(),
  }),
  useCreateGsmWaybillMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    reset: vi.fn(),
  }),
}));

describe("GsmPeriodView", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("loads strip after selecting period and wires generate to API", async () => {
    mockUseWaybills.mockImplementation((params: { vehicleId?: number } | null) => ({
      isLoading: false,
      error: null,
      data: params ? [WAYBILL] : [],
      refetch: mockRefetch,
    }));

    const result: WaybillGenerateResult = {
      waybills: [WAYBILL],
      warnings: ["weekend_anchor"],
      days_created: 1,
      problematic_days: [],
      manual_days: 0,
    };
    mockGenerate.mockResolvedValue(result);

    render(<GsmPeriodView />);

    expect(screen.getByText(/Выберите машину и период/i)).toBeInTheDocument();

    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });

    expect(await screen.findByLabelText("Период Geely Monjaro")).toBeInTheDocument();
    expect(screen.getByText("Якорь")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Ручной ПЛ" })).toBeInTheDocument();

    fireEvent.click(screen.getByLabelText("День 2025-04-03"));
    expect(await screen.findByText(/Правка дня 2025-04-03/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Закрыть" }));

    fireEvent.change(screen.getByLabelText("Стартовый остаток бака"), { target: { value: "20" } });
    fireEvent.change(screen.getByLabelText("Стартовый одометр"), { target: { value: "10000" } });
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать" }));

    await waitFor(() => {
      expect(mockGenerate).toHaveBeenCalledWith({
        vehicle_id: 1,
        period_from: "2025-04-01",
        period_to: "2025-04-30",
        force: false,
        fuel_start: 20,
        odometer_start: 10000,
      });
    });

    expect(await screen.findByText(/Создано 1 дней, 0 требуют ручной доработки/)).toBeInTheDocument();

    const periodBadge = screen.getByRole("button", { name: /Предупреждение периода: Выходной/ });
    fireEvent.click(periodBadge);
    expect(screen.getByRole("status")).toHaveTextContent(/выходной или праздник/i);
  });

  it("shows partial generation summary and problematic dates after generate 200", async () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [],
      refetch: mockRefetch,
    });
    mockGenerate.mockResolvedValue({
      waybills: [WAYBILL],
      warnings: ["weekend_anchor"],
      days_created: 10,
      problematic_days: [
        {
          date: "2025-04-07",
          reason: "manual_intervention",
          detail: "не удалось сжечь 51.2 л",
          fuel_before: 40.1,
          fuel_to_issue: 54.57,
          tank_volume: 60,
        },
        {
          date: "2025-04-08",
          reason: "manual_intervention",
          detail: "коридор бака",
          fuel_before: 30,
          fuel_to_issue: 40,
          tank_volume: 60,
        },
      ],
      manual_days: 2,
    } satisfies WaybillGenerateResult);

    render(<GsmPeriodView />);
    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать" }));

    expect(await screen.findByText(/Создано 10 дней, 2 требуют ручной доработки/)).toBeInTheDocument();
    const dates = screen.getByLabelText("Дни ручной доработки");
    expect(dates).toHaveTextContent("2025-04-07");
    expect(dates).toHaveTextContent("2025-04-08");
    expect(screen.queryByRole("button", { name: /Предупреждение периода: Нерешаемо/ })).not.toBeInTheDocument();
  });

  it("treats 422 as a configuration error, not an unsolvable period", async () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [],
      refetch: mockRefetch,
    });
    mockGenerate.mockRejectedValue(
      new ApiError("routes", 422, "vehicle has no routes in library", "gsm_routes_required"),
    );

    render(<GsmPeriodView />);
    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });
    fireEvent.click(screen.getByRole("button", { name: "Сгенерировать" }));

    expect(await screen.findByText(/нет маршрутов/i)).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: /Предупреждение периода: Нерешаемо/ })).not.toBeInTheDocument();
    expect(screen.queryByText(/требуют ручной доработки/)).not.toBeInTheDocument();
  });

  it("disables export when the period has no waybills", () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [],
      refetch: mockRefetch,
    });

    render(<GsmPeriodView />);
    expect(screen.getByRole("button", { name: "Экспорт zip" })).toBeDisabled();

    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });

    expect(screen.getByRole("button", { name: "Экспорт zip" })).toBeDisabled();
    expect(screen.getByText("Нет путевых листов за период.")).toBeInTheDocument();
  });

  it("blocks export when a day needs manual intervention", () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [{ ...WAYBILL, warnings: ["manual_intervention"] }],
      refetch: mockRefetch,
    });

    render(<GsmPeriodView />);
    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });

    expect(screen.getByRole("button", { name: "Экспорт zip" })).toBeDisabled();
    expect(screen.getByText(/Исправьте дни ручной доработки: 2025-04-03/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Экспорт zip" }));
    expect(mockExport).not.toHaveBeenCalled();
  });

  it("exports clean drafts immediately", async () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [{ ...WAYBILL, warnings: [] }],
      refetch: mockRefetch,
    });
    mockExport.mockResolvedValue({ blob: new Blob(), filename: "x.zip" });

    render(<GsmPeriodView />);
    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });

    fireEvent.click(screen.getByRole("button", { name: "Экспорт zip" }));
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();

    await waitFor(() => {
      expect(mockExport).toHaveBeenCalledWith({
        vehicle_ids: [1],
        from: "2025-04-01",
        to: "2025-04-30",
      });
    });
    expect(await screen.findByText(/Скачан zip с бланками/)).toBeInTheDocument();
  });

  it("asks to confirm yellow warnings before export", async () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [WAYBILL],
      refetch: mockRefetch,
    });
    mockExport.mockResolvedValue({ blob: new Blob(), filename: "x.zip" });

    render(<GsmPeriodView />);
    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });

    fireEvent.click(screen.getByRole("button", { name: "Экспорт zip" }));
    expect(mockExport).not.toHaveBeenCalled();
    expect(await screen.findByRole("dialog")).toHaveTextContent(/предупреждения/i);

    fireEvent.click(screen.getByRole("button", { name: "Скачать" }));
    await waitFor(() => {
      expect(mockExport).toHaveBeenCalledTimes(1);
    });
  });

  it("asks to confirm re-export of already exported days", async () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [{ ...WAYBILL, status: "exported", warnings: [] }],
      refetch: mockRefetch,
    });
    mockExport.mockResolvedValue({ blob: new Blob(), filename: "x.zip" });

    render(<GsmPeriodView />);
    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });

    fireEvent.click(screen.getByRole("button", { name: "Экспорт zip" }));
    expect(await screen.findByRole("dialog")).toHaveTextContent(/уже экспортировался/i);
    fireEvent.click(screen.getByRole("button", { name: "Отмена" }));
    expect(mockExport).not.toHaveBeenCalled();
  });

  it("shows LibreOffice hint when export fails", async () => {
    mockUseWaybills.mockReturnValue({
      isLoading: false,
      error: null,
      data: [{ ...WAYBILL, warnings: [] }],
      refetch: mockRefetch,
    });
    mockExport.mockRejectedValue(
      new ApiError(
        "soffice",
        500,
        "LibreOffice (soffice) is not installed or not on PATH",
        "gsm_export_soffice_missing",
      ),
    );

    render(<GsmPeriodView />);
    fireEvent.change(screen.getByLabelText("Машина"), { target: { value: "1" } });
    fireEvent.change(screen.getByLabelText("Период с"), { target: { value: "2025-04-01" } });
    fireEvent.change(screen.getByLabelText("Период по"), { target: { value: "2025-04-30" } });
    fireEvent.click(screen.getByRole("button", { name: "Экспорт zip" }));

    expect(await screen.findByText(/LibreOffice \(soffice\)/)).toBeInTheDocument();
  });
});
