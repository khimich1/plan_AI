import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { OfferDetailsDrawer } from "@/features/commercial-archive/components/OfferDetailsDrawer";
import type { ArchiveOfferDetails, KpReadinessSummary } from "@/features/commercial-archive/types/archive";

const mockUseArchiveOfferQuery = vi.fn();

vi.mock("@/features/commercial-archive/hooks/useArchiveQueries", () => ({
  useArchiveOfferQuery: (...args: unknown[]) => mockUseArchiveOfferQuery(...args),
  useArchiveDocumentMutation: () => ({ mutate: vi.fn(), isPending: false, isError: false, error: null }),
  useUpdateDiscountMutation: () => ({ mutateAsync: vi.fn(), isPending: false, isError: false, error: null }),
  useUpdateLogisticsCostMutation: () => ({ mutateAsync: vi.fn(), isPending: false, isError: false, error: null }),
}));

vi.mock("@/features/delivery-schedule/hooks/useDeliveryScheduleQueries", () => ({
  useDeliveryScheduleQuery: () => ({
    data: null,
    isPending: false,
    isError: false,
    error: null,
  }),
  usePutDeliveryScheduleMutation: () => ({
    mutateAsync: vi.fn(),
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
}));

vi.mock("@/features/delivery-schedule/components/DeliveryScheduleDialog", () => ({
  DeliveryScheduleDialog: () => null,
}));

vi.mock("@/features/commercial-archive/components/KpReadinessBlock", () => ({
  KpReadinessBlock: () => <div data-testid="kp-readiness-block">KpReadinessBlock</div>,
}));

vi.mock("@/features/commercial-archive/components/DeleteConfirmDialog", () => ({
  DeleteConfirmDialog: () => null,
}));

vi.mock("@/features/commercial-archive/components/MoveToProductionDialog", () => ({
  MoveToProductionDialog: () => null,
}));

function makeReadiness(): KpReadinessSummary {
  return {
    completion_percentage: 72,
    sgp_progress: { n: 14, m: 20 },
    issuable_qty: 14,
    in_production_qty: 6,
    summary_text: "14 из 20 шт на складе, 6 в производстве. Можно выдать 14 шт.",
    client_copy_text: "Здравствуйте! По вашему заказу №42: 14 из 20 шт уже на складе.",
    steps: [
      { id: "kp", label: "КП", state: "done" },
      { id: "production", label: "Производство", state: "active", hint: "72%" },
      { id: "sgp", label: "СГП", state: "active", hint: "14/20" },
      { id: "release", label: "Выдача", state: "disabled" },
      { id: "closed", label: "Закрыто", state: "disabled" },
    ],
    release_note: "Выдача с СГП — в следующем обновлении",
  };
}

function makeOffer(
  status: string,
  readiness: KpReadinessSummary | null = null,
  overrides: Partial<ArchiveOfferDetails> = {},
): ArchiveOfferDetails {
  return {
    kp_id: 42,
    creation_date: "01.03.2026",
    customer_name: "ООО Тест",
    manager_name: "Иван Иванов",
    status,
    execution_terms: null,
    delivery_conditions: null,
    payment_conditions: null,
    finance: {
      subtotal: 1000,
      vat_amount: 220,
      total_amount: 1220,
      discount_percent: 5,
    },
    logistics_cost: 0,
    total_cargo_weight_kg: 0,
    delivery_service_total_rub: 0,
    plates: [],
    completion_percentage: null,
    readiness,
    ...overrides,
  };
}

function makePileOffer(status = "в архиве"): ArchiveOfferDetails {
  return makeOffer(status, null, {
    product_type: "piles",
    piles: [
      {
        position_number: 1,
        mark: "С80.30-8",
        concrete_grade: "B25",
        qty: 10,
        unit_price: 5000,
        discounted_price: 4750,
      },
    ],
  });
}

describe("OfferDetailsDrawer readiness visibility", () => {
  beforeEach(() => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: undefined,
      isPending: false,
      isError: false,
      error: null,
    });
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it.each(["в архиве", "выполнено"])("does not render KpReadinessBlock when status is %s", (status) => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer(status, makeReadiness()),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.queryByTestId("kp-readiness-block")).not.toBeInTheDocument();
  });

  it("renders KpReadinessBlock when status is в работе and readiness is present", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе", makeReadiness()),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByTestId("kp-readiness-block")).toBeInTheDocument();
  });

  it("enables delivery schedule button for в работе and архиве (read-only for archive)", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе", makeReadiness()),
      isPending: false,
      isError: false,
      error: null,
    });
    const { rerender } = render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);
    expect(screen.getByRole("button", { name: "График поставки" })).toBeEnabled();
    expect(screen.getByRole("button", { name: "График поставки" })).toHaveAttribute(
      "title",
      "Редактирование графика поставки",
    );

    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", makeReadiness()),
      isPending: false,
      isError: false,
      error: null,
    });
    rerender(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);
    expect(screen.getByRole("button", { name: "График поставки" })).toBeEnabled();
    expect(screen.getByRole("button", { name: "График поставки" })).toHaveAttribute(
      "title",
      "Просмотр графика поставки",
    );
  });
});

describe("OfferDetailsDrawer pile offers", () => {
  beforeEach(() => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: undefined,
      isPending: false,
      isError: false,
      error: null,
    });
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("does not render KpReadinessBlock for pile offers in production", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makePileOffer("в работе"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.queryByTestId("kp-readiness-block")).not.toBeInTheDocument();
  });

  it("renders pile table columns and row data", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makePileOffer(),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByText("Марка")).toBeInTheDocument();
    expect(screen.getByText("Класс")).toBeInTheDocument();
    expect(screen.getByText("С80.30-8")).toBeInTheDocument();
    expect(screen.getByText("B25")).toBeInTheDocument();
  });

  it("hides schema button for pile offers", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makePileOffer(),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.queryByRole("button", { name: /Схема/i })).not.toBeInTheDocument();
  });

  it("disables move to production for archived pile offers", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makePileOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    const moveButton = screen.getByRole("button", { name: /В производство/i });
    expect(moveButton).toBeDisabled();
    expect(moveButton).toHaveAttribute("title", "скоро");
  });
});
