import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ShipmentDrawer } from "@/features/logistics/components/ShipmentDrawer";
import type {
  PileCatalogEntry,
  ProposeResponse,
  ShipmentDetails,
} from "@/features/logistics/types/logistics";

const mockUseShipmentQuery = vi.fn();
const mockPropose = vi.fn();
const mockConfirm = vi.fn();

const PILE_CATALOG: PileCatalogEntry[] = [
  { id: 1, mark: "С60.30", length_m: 6, section_mm: 300, volume_m3: 0.54, weight_kg: 1380 },
];

vi.mock("@/features/logistics/hooks/useLogisticsQueries", () => ({
  useShipmentQuery: (...args: unknown[]) => mockUseShipmentQuery(...args),
  useUpdateShipmentMutation: () => ({ mutateAsync: vi.fn(), isPending: false }),
  useCompleteShipmentMutation: () => ({ mutateAsync: vi.fn(), isPending: false }),
  useCancelShipmentMutation: () => ({ mutateAsync: vi.fn(), isPending: false }),
  useShipmentSheetMutation: () => ({ mutateAsync: vi.fn(), isPending: false }),
  useProposeMutation: () => ({ mutateAsync: mockPropose, isPending: false }),
  useConfirmItemsMutation: () => ({ mutateAsync: mockConfirm, isPending: false }),
  usePileCatalogQuery: () => ({ data: PILE_CATALOG, isLoading: false }),
  useCarriersQuery: () => ({ data: [], isLoading: false }),
}));

const SHIPMENT: ShipmentDetails = {
  id: 5,
  shipment_date: "2026-08-05",
  delivery_type: "delivery",
  status: "in_work",
  attention: 0,
  attention_comment: null,
  carrier_id: null,
  carrier_name: null,
  driver_name: null,
  vehicle_text: null,
  vehicle_class: "t20",
  proxy_no: null,
  upd_no: null,
  freight_request_no: null,
  planned_cost: null,
  total_weight_kg: null,
  completed_at: null,
  orders: [{ id: 1, kp_id: 10, ya_order_no: "ЯР-0001467", customer_name: "ООО Ромашка" }],
  items: [],
  available_by_kp: [
    {
      kp_id: 10,
      plates: [
        {
          completed_plate_id: 77,
          plate_name: "ПБ 60-12-8",
          length_m: 6,
          width_m: 1.2,
          load_class: 800,
          available_qty: 20,
          unit_weight_kg: 2700,
        },
      ],
    },
  ],
};

const PROPOSE_RESPONSE: ProposeResponse = {
  items: [
    {
      item_type: "plate",
      completed_plate_id: 77,
      kp_id: 10,
      plate_name: "ПБ 60-12-8",
      length_m: 6,
      width_m: 1.2,
      load_class: 800,
      qty: 10,
      available_qty: 20,
      unit_weight_kg: 2700,
      weight_kg: 27_000,
    },
  ],
  not_fit: [
    {
      item_type: "plate",
      completed_plate_id: 78,
      kp_id: 10,
      plate_name: "ПБ 45-12-8",
      length_m: 4.5,
      width_m: 1.2,
      load_class: 800,
      qty: 5,
      available_qty: 8,
      unit_weight_kg: 2030,
      weight_kg: 10_150,
      reason_code: "weight_limit",
      reason_text: "Превышен лимит веса класса ТС",
    },
  ],
  order_remainder: [
    {
      completed_plate_id: 78,
      kp_id: 10,
      plate_name: "ПБ 45-12-8",
      qty_remaining: 5,
    },
  ],
  warnings: [],
  total_weight_kg: 27_000,
  overload: true,
  vehicle_class: "t20",
  vehicle_class_limits_kg: { t20: 19_800, t30plus: 30_000 },
  layout: {
    body_length_m: 13.2,
    body_used_m: 6.0,
    stacks: [
      {
        index: 1,
        marking_length_m: 6.0,
        tiers: [
          {
            index: 1,
            units: [
              { completed_plate_id: 77, kp_id: 10, plate_name: "ПБ 60-12-8", width_m: 1.2 },
              { completed_plate_id: 77, kp_id: 10, plate_name: "ПБ 60-12-8", width_m: 1.2 },
            ],
          },
        ],
      },
    ],
    loading_steps: [
      { step: 1, stack_index: 1, tier_index: 1, description: "ПБ 60-12-8 ×2" },
    ],
  },
};

describe("ShipmentDrawer propose → confirm", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  const renderDrawer = () => {
    mockUseShipmentQuery.mockReturnValue({
      data: SHIPMENT,
      isLoading: false,
      isError: false,
      error: null,
    });
    mockPropose.mockResolvedValue(PROPOSE_RESPONSE);
    mockConfirm.mockResolvedValue([]);
    return render(<ShipmentDrawer shipmentId={5} onClose={() => {}} />);
  };

  it("proposes items, shows not_fit and overload warning, confirms edited qty", async () => {
    renderDrawer();

    fireEvent.click(screen.getByRole("button", { name: "Предложить состав" }));

    expect(await screen.findByText(/Предложенный состав подставлен/)).toBeInTheDocument();
    expect(screen.getByText("Не влезло в лимит класса ТС (1 поз.)")).toBeInTheDocument();
    expect(screen.queryByText(/Превышен лимит веса класса ТС/)).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Добавить всё равно" })).not.toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: /Остаток по заказу \(на следующий рейс\) \(1 поз\.\)/ }),
    ).toBeInTheDocument();
    expect(screen.queryByText(/ПБ 45-12-8 · 5 шт \(КП №10\)/)).not.toBeInTheDocument();

    fireEvent.click(
      screen.getByRole("button", { name: /Остаток по заказу \(на следующий рейс\) \(1 поз\.\)/ }),
    );
    expect(screen.getByText(/ПБ 45-12-8 · 5 шт \(КП №10\)/)).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: /Не влезло в лимит класса ТС \(1 поз\.\)/ }));
    expect(screen.getByText(/Превышен лимит веса класса ТС/)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Добавить всё равно" })).toBeInTheDocument();

    // Перегруз виден, но не блокирует утверждение
    expect(screen.getByRole("alert")).toHaveTextContent("Перегруз");

    const qtyInput = screen.getByDisplayValue("10");
    fireEvent.change(qtyInput, { target: { value: "3" } });

    fireEvent.click(screen.getByRole("button", { name: "Утвердить состав" }));

    expect(mockConfirm).toHaveBeenCalledWith([
      expect.objectContaining({
        item_type: "plate",
        completed_plate_id: 77,
        kp_id: 10,
        qty: 3,
        sort_order: 0,
      }),
    ]);
  });

  it("shows layout after propose and hides it on manual edit", async () => {
    renderDrawer();

    fireEvent.click(screen.getByRole("button", { name: "Предложить состав" }));
    expect(await screen.findByText(/Укладка в кузов/)).toBeInTheDocument();
    expect(screen.getByText("Порядок погрузки")).toBeInTheDocument();
    expect(screen.getByText(/Штабель 1, ярус 1: ПБ 60-12-8 ×2/)).toBeInTheDocument();

    const qtyInput = screen.getByDisplayValue("10");
    fireEvent.change(qtyInput, { target: { value: "3" } });
    expect(screen.queryByText(/Укладка в кузов/)).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Предложить состав" }));
    expect(await screen.findByText(/Укладка в кузов/)).toBeInTheDocument();
  });

  it("adds free row with pile-catalog auto weight and confirms it", async () => {
    renderDrawer();

    fireEvent.click(screen.getByRole("button", { name: "+ Свободная строка" }));

    const markInput = screen.getByPlaceholderText("Марка (С60.30)");
    fireEvent.change(markInput, { target: { value: "С60.30" } });

    expect(screen.getByText(/автовес: 1 380 кг \/ шт/)).toBeInTheDocument();
    expect(screen.getByText(/Σ веса: 1 380 кг/)).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Утвердить состав" }));

    expect(mockConfirm).toHaveBeenCalledWith([
      expect.objectContaining({
        item_type: "free",
        mark: "С60.30",
        qty: 1,
        weight_kg: undefined,
        sort_order: 0,
      }),
    ]);
  });

  it("adds plate row from available_by_kp picker", async () => {
    renderDrawer();

    fireEvent.click(screen.getByRole("button", { name: "+ Плита со СГП" }));
    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));

    const row = screen.getByText("ПБ 60-12-8").closest("tr");
    expect(row).not.toBeNull();
    expect(screen.getByText(/Σ веса: 2 700 кг/)).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Утвердить состав" }));

    expect(mockConfirm).toHaveBeenCalledWith([
      expect.objectContaining({
        item_type: "plate",
        completed_plate_id: 77,
        qty: 1,
        sort_order: 0,
      }),
    ]);
  });
});
