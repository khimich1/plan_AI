import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { ShipmentRegistryRow } from "@/features/logistics/types/logistics";
import { LogisticsRegistryView } from "@/features/logistics/components/LogisticsRegistryView";

const mockUseShipmentsQuery = vi.fn();
const dialogPropsSpy = vi.fn();
const drawerPropsSpy = vi.fn();

vi.mock("@/features/logistics/hooks/useLogisticsQueries", () => ({
  useShipmentsQuery: (...args: unknown[]) => mockUseShipmentsQuery(...args),
  useCarriersQuery: () => ({ data: [], isLoading: false, isError: false, error: null }),
}));

vi.mock("@/features/logistics/components/ShipmentDrawer", () => ({
  ShipmentDrawer: (props: { shipmentId: number | null }) => {
    drawerPropsSpy(props);
    return props.shipmentId != null ? (
      <div data-testid="shipment-drawer">drawer-{props.shipmentId}</div>
    ) : null;
  },
}));

vi.mock("@/features/logistics/components/CreateShipmentDialog", () => ({
  CreateShipmentDialog: (props: {
    open: boolean;
    sourceShipmentId?: number;
    initialDeliveryType?: string;
    onCreated: (id: number) => void;
  }) => {
    dialogPropsSpy(props);
    return props.open ? (
      <div data-testid="create-dialog">
        <span>source:{props.sourceShipmentId ?? "none"}</span>
        <span>dtype:{props.initialDeliveryType ?? "none"}</span>
        <button type="button" onClick={() => props.onCreated(99)}>
          mock-created
        </button>
      </div>
    ) : null;
  },
}));

const makeRow = (overrides: Partial<ShipmentRegistryRow> & Pick<ShipmentRegistryRow, "id">): ShipmentRegistryRow => ({
  shipment_date: "2026-08-01",
  delivery_type: "delivery",
  status: "in_work",
  attention: 0,
  attention_comment: null,
  carrier_name: "ООО ТрансЛогистик",
  proxy_no: null,
  driver_name: "Иванов И.И.",
  vehicle_text: "Volvo FH / а123бв77",
  upd_no: "101",
  planned_cost: 45_000,
  total_weight_kg: 18_400,
  orders: [{ kp_id: 154, ya_order_no: "ЯР-0001467", customer_name: "ООО Ромашка" }],
  ...overrides,
});

const ROWS: ShipmentRegistryRow[] = [
  makeRow({ id: 1, attention: 1, attention_comment: "Работа крана!" }),
  makeRow({
    id: 2,
    shipment_date: "2026-08-02",
    delivery_type: "pickup",
    status: "done",
    carrier_name: null,
    proxy_no: "77-12",
    upd_no: null,
    orders: [{ kp_id: 160, ya_order_no: "ЯР-0002000", customer_name: "ИП Василёк" }],
  }),
];

const lastFilters = (): Record<string, unknown> => {
  const calls = mockUseShipmentsQuery.mock.calls;
  return calls[calls.length - 1][0] as Record<string, unknown>;
};

describe("LogisticsRegistryView", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  const renderView = () => {
    mockUseShipmentsQuery.mockReturnValue({
      data: ROWS,
      isLoading: false,
      isError: false,
      error: null,
    });
    return render(<LogisticsRegistryView />);
  };

  it("renders registry rows with statuses and attention icon", () => {
    renderView();
    expect(screen.getByText("ЯР-0001467")).toBeInTheDocument();
    expect(screen.getByText("ЯР-0002000")).toBeInTheDocument();
    expect(screen.getByText("В работе")).toBeInTheDocument();
    expect(screen.getByText("Обработано")).toBeInTheDocument();
    expect(screen.getByRole("img", { name: "Внимание" })).toBeInTheDocument();
    expect(screen.getByText("дов. 77-12")).toBeInTheDocument();
  });

  it("passes no_upd flag to query when «Без УПД» checked", () => {
    renderView();
    fireEvent.click(screen.getByRole("checkbox", { name: "Без УПД" }));
    expect(lastFilters().no_upd).toBe(true);
  });

  it("passes attention flag to query when «Внимание» checked", () => {
    renderView();
    fireEvent.click(screen.getByRole("checkbox", { name: "Внимание" }));
    expect(lastFilters().attention).toBe(true);
  });

  it("maps numeric order query to kp_id server filter", () => {
    renderView();
    fireEvent.change(screen.getByPlaceholderText("например 154 или ЯР-0001"), {
      target: { value: "154" },
    });
    expect(lastFilters().kp_id).toBe(154);
  });

  it("filters rows client-side by ЯР/customer for non-numeric order query", () => {
    renderView();
    fireEvent.change(screen.getByPlaceholderText("например 154 или ЯР-0001"), {
      target: { value: "василёк" },
    });
    expect(lastFilters().kp_id).toBeUndefined();
    expect(screen.queryByText("ЯР-0001467")).not.toBeInTheDocument();
    expect(screen.getByText("ЯР-0002000")).toBeInTheDocument();
  });

  it("passes date range and delivery type to query", () => {
    renderView();
    fireEvent.change(screen.getByLabelText("Дата с"), { target: { value: "2026-08-01" } });
    fireEvent.change(screen.getByLabelText("Дата по"), { target: { value: "2026-08-31" } });
    fireEvent.change(screen.getByRole("combobox", { name: "Тип" }), {
      target: { value: "pickup" },
    });
    expect(lastFilters().date_from).toBe("2026-08-01");
    expect(lastFilters().date_to).toBe("2026-08-31");
    expect(lastFilters().delivery_type).toBe("pickup");
  });

  it("resets filters via «Сбросить»", () => {
    renderView();
    fireEvent.click(screen.getByRole("checkbox", { name: "Без УПД" }));
    expect(lastFilters().no_upd).toBe(true);
    fireEvent.click(screen.getByRole("button", { name: "Сбросить" }));
    expect(lastFilters().no_upd).toBeUndefined();
  });

  it("«На основе» opens create dialog with source id and does not open source drawer", () => {
    renderView();
    const reuseButtons = screen.getAllByRole("button", { name: "На основе" });
    expect(reuseButtons[0]).toHaveAttribute("title", "Создать на основе");

    fireEvent.click(reuseButtons[1]); // row id=2, pickup

    expect(screen.getByTestId("create-dialog")).toBeInTheDocument();
    expect(screen.getByText("source:2")).toBeInTheDocument();
    expect(screen.getByText("dtype:pickup")).toBeInTheDocument();
    expect(screen.queryByTestId("shipment-drawer")).not.toBeInTheDocument();
  });

  it("after reuse success opens drawer for the new shipment id", () => {
    renderView();
    fireEvent.click(screen.getAllByRole("button", { name: "На основе" })[0]);
    fireEvent.click(screen.getByRole("button", { name: "mock-created" }));

    expect(screen.getByTestId("shipment-drawer")).toHaveTextContent("drawer-99");
    expect(screen.queryByTestId("create-dialog")).not.toBeInTheDocument();
  });

  it("row click still opens the source drawer", () => {
    renderView();
    fireEvent.click(screen.getByText("ЯР-0001467"));
    expect(screen.getByTestId("shipment-drawer")).toHaveTextContent("drawer-1");
  });
});
