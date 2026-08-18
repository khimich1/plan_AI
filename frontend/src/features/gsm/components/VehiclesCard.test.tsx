import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { VehiclesCard } from "@/features/gsm/components/VehiclesCard";
import { ApiError } from "@/shared/lib/apiError";
import type { GsmVehicle } from "@/features/gsm/types/gsm";

const mockPatch = vi.fn();
const mockUseVehicles = vi.fn();

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useGsmVehiclesQuery: (...args: unknown[]) => mockUseVehicles(...args),
  usePatchVehicleMutation: () => ({ mutateAsync: mockPatch, isPending: false }),
}));

const VEHICLE: GsmVehicle = {
  id: 1,
  name: "Geely Monjaro",
  plate_number: "A123BC77",
  tank_volume_liters: 60,
  norm_summer: 12.5,
  norm_winter: 14.1,
  primary_driver_id: null,
  is_active: true,
};

describe("VehiclesCard", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("edits norms and tank via patch mutation", async () => {
    mockUseVehicles.mockReturnValue({ data: [VEHICLE], isLoading: false, error: null });
    mockPatch.mockResolvedValue({ ...VEHICLE, tank_volume_liters: 55, norm_summer: 11, norm_winter: 13 });
    render(<VehiclesCard />);

    fireEvent.click(screen.getByRole("button", { name: "Нормы / бак" }));
    fireEvent.change(screen.getByLabelText("Бак Geely Monjaro"), { target: { value: "55" } });
    fireEvent.change(screen.getByLabelText("Норма лето Geely Monjaro"), { target: { value: "11" } });
    fireEvent.change(screen.getByLabelText("Норма зима Geely Monjaro"), { target: { value: "13" } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить" }));

    await waitFor(() => {
      expect(mockPatch).toHaveBeenCalledWith({
        id: 1,
        payload: { tank_volume_liters: 55, norm_summer: 11, norm_winter: 13 },
      });
    });
    expect(await screen.findByText(/Нормы и бак сохранены/)).toBeInTheDocument();
  });

  it("shows human-readable validation error from API", async () => {
    mockUseVehicles.mockReturnValue({ data: [VEHICLE], isLoading: false, error: null });
    mockPatch.mockRejectedValue(
      new ApiError("bad", 422, "tank_volume_liters must be > 0", "gsm_validation"),
    );
    render(<VehiclesCard />);

    fireEvent.click(screen.getByRole("button", { name: "Нормы / бак" }));
    fireEvent.click(screen.getByRole("button", { name: "Сохранить" }));

    expect(await screen.findByText(/Объём бака должен быть больше 0/)).toBeInTheDocument();
  });
});
