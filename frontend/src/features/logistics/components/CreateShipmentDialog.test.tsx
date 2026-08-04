import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CreateShipmentDialog } from "@/features/logistics/components/CreateShipmentDialog";

const mutateCreate = vi.fn();
const mutateReuse = vi.fn();

vi.mock("@/features/logistics/hooks/useLogisticsQueries", () => ({
  useCreateShipmentMutation: () => ({
    mutateAsync: mutateCreate,
    isPending: false,
  }),
  useReuseTransportMutation: () => ({
    mutateAsync: mutateReuse,
    isPending: false,
  }),
}));

const searchKp = vi.fn();

vi.mock("@/features/logistics/api/logisticsApi", () => ({
  logisticsApi: {
    searchKp: (...args: unknown[]) => searchKp(...args),
  },
}));

describe("CreateShipmentDialog", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  beforeEach(() => {
    mutateCreate.mockResolvedValue({ id: 11 });
    mutateReuse.mockResolvedValue({ id: 22 });
    searchKp.mockResolvedValue({ mode: "number", items: [], total: 0, truncated: false });
  });

  it("searches KP via logisticsApi and shows ACL empty message", async () => {
    searchKp.mockResolvedValue({ mode: "number", items: [], total: 0, truncated: false });
    render(<CreateShipmentDialog open onClose={vi.fn()} onCreated={vi.fn()} />);

    fireEvent.change(screen.getByPlaceholderText("Номер КП или заказчик"), {
      target: { value: "154" },
    });
    fireEvent.click(screen.getByRole("button", { name: "Найти" }));

    await waitFor(() => {
      expect(searchKp).toHaveBeenCalledWith({ kpId: 154 });
    });
    expect(
      screen.getByText(
        "КП не найдено. Для поиска КП должно быть в статусе «в работе» или «На СГП».",
      ),
    ).toBeInTheDocument();
  });

  it("submits createShipment when sourceShipmentId is absent", async () => {
    const onCreated = vi.fn();
    render(<CreateShipmentDialog open onClose={vi.fn()} onCreated={onCreated} />);

    fireEvent.change(screen.getByPlaceholderText("Номер КП или заказчик"), {
      target: { value: "154" },
    });
    fireEvent.click(screen.getByRole("button", { name: "Добавить по номеру" }));
    fireEvent.click(screen.getByRole("button", { name: "Создать рейс" }));

    await waitFor(() => {
      expect(mutateCreate).toHaveBeenCalledWith(
        expect.objectContaining({
          delivery_type: "delivery",
          kp_ids: [154],
        }),
      );
    });
    expect(mutateReuse).not.toHaveBeenCalled();
    expect(onCreated).toHaveBeenCalledWith(11);
  });

  it("prefills delivery_type and submits reuseTransport when source is set", async () => {
    const onCreated = vi.fn();
    render(
      <CreateShipmentDialog
        open
        onClose={vi.fn()}
        onCreated={onCreated}
        sourceShipmentId={5}
        initialDeliveryType="pickup"
      />,
    );

    expect(screen.getByText("Новый рейс на основе")).toBeInTheDocument();
    expect(screen.getByRole("combobox")).toHaveValue("pickup");

    fireEvent.change(screen.getByPlaceholderText("Номер КП или заказчик"), {
      target: { value: "160" },
    });
    fireEvent.click(screen.getByRole("button", { name: "Добавить по номеру" }));
    fireEvent.click(screen.getByRole("button", { name: "Создать на основе" }));

    await waitFor(() => {
      expect(mutateReuse).toHaveBeenCalledWith({
        sourceId: 5,
        payload: expect.objectContaining({
          delivery_type: "pickup",
          kp_ids: [160],
        }),
      });
    });
    expect(mutateCreate).not.toHaveBeenCalled();
    expect(onCreated).toHaveBeenCalledWith(22);
  });
});
