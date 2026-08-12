import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { OfferDetailsDrawer } from "@/features/commercial-archive/components/OfferDetailsDrawer";
import type { ArchiveOfferDetails, KpReadinessSummary } from "@/features/commercial-archive/types/archive";

const mockUseArchiveOfferQuery = vi.fn();
const mockResume = vi.fn();
const mockNavigate = vi.fn();
const mockDispatch = vi.fn();

vi.mock("react-router", async () => {
  const actual = await vi.importActual<typeof import("react-router")>("react-router");
  return {
    ...actual,
    useNavigate: () => mockNavigate,
  };
});

vi.mock("@/features/commercial-offer/store/wizardDraftStore", () => ({
  useWizardDraftStore: () => ({
    state: { draftId: null },
    dispatch: mockDispatch,
  }),
}));

vi.mock("@/features/commercial-archive/hooks/useArchiveQueries", () => ({
  useArchiveOfferQuery: (...args: unknown[]) => mockUseArchiveOfferQuery(...args),
  useArchiveDocumentMutation: () => ({ mutate: vi.fn(), isPending: false, isError: false, error: null }),
  useUpdateDiscountMutation: () => ({ mutateAsync: vi.fn(), isPending: false, isError: false, error: null }),
  useUpdateLogisticsCostMutation: () => ({ mutateAsync: vi.fn(), isPending: false, isError: false, error: null }),
}));

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    resume: (...args: unknown[]) => mockResume(...args),
    buildDocumentUrl: (kpId: number, kind: string) =>
      `/api/v1/commercial/archive/${kpId}/files/${kind}`,
  },
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

function makeResumeDraft(draftId = "draft-resume-42") {
  return {
    draft_id: draftId,
    order: {},
    optimization: { total_plates: 0, total_cost: 0 },
    order_data: [],
    metadata: {
      client_name: "ООО Тест",
      discount_percent: 5,
      manager_id: 1,
      conditions_mode: "standard" as const,
      delivery_conditions: "",
      payment_conditions: "",
      execution_terms: "",
      logistics_cost: 0,
      product_type: "plates" as const,
      resume_kp_id: 42,
    },
    wizard_state: { current_step: "result" },
    files: [],
    saved_offer: { kp_id: 42, status: "в работе", execution_terms: "" },
    totals: {},
    offer_identity: { title: "КП", subtitle: "" },
  };
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

describe("OfferDetailsDrawer MNA-602 — append CTA", () => {
  beforeEach(() => {
    mockResume.mockReset();
    mockNavigate.mockReset();
    mockDispatch.mockReset();
    mockResume.mockResolvedValue(makeResumeDraft());
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

  it("renders CTA «Добавить другое наименование» when status is в работе", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(
      screen.getByRole("button", { name: "Добавить другое наименование" }),
    ).toBeInTheDocument();
  });

  it.each(["в архиве", "выполнено", "На СГП", "отклонено", "в ожидании"])(
    "does not render append CTA when status is %s",
    (status) => {
      mockUseArchiveOfferQuery.mockReturnValue({
        data: makeOffer(status),
        isPending: false,
        isError: false,
        error: null,
      });

      render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

      expect(
        screen.queryByRole("button", { name: "Добавить другое наименование" }),
      ).not.toBeInTheDocument();
    },
  );

  it("resumes draft, hydrates store, starts append cycle, and navigates to wizard", async () => {
    const onClose = vi.fn();
    const draft = makeResumeDraft("draft-resume-42");
    mockResume.mockResolvedValue(draft);
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={onClose} />);

    fireEvent.click(screen.getByRole("button", { name: "Добавить другое наименование" }));

    await waitFor(() => {
      expect(mockResume).toHaveBeenCalledWith(42);
    });

    await waitFor(() => {
      expect(mockDispatch).toHaveBeenCalledWith({ type: "hydrate-draft", payload: draft });
      expect(mockDispatch).toHaveBeenCalledWith({ type: "start-append-cycle" });
      expect(mockNavigate).toHaveBeenCalledWith("/new?draft=draft-resume-42");
      expect(onClose).toHaveBeenCalled();
    });

    const hydrateIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "hydrate-draft",
    );
    const appendIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "start-append-cycle",
    );
    expect(hydrateIndex).toBeGreaterThanOrEqual(0);
    expect(appendIndex).toBeGreaterThan(hydrateIndex);
  });

  it("disables CTA while resume is pending", async () => {
    let resolveResume: (value: unknown) => void = () => undefined;
    mockResume.mockReturnValue(
      new Promise((resolve) => {
        resolveResume = resolve;
      }),
    );
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    fireEvent.click(screen.getByRole("button", { name: "Добавить другое наименование" }));

    await waitFor(() => {
      expect(screen.getByRole("button", { name: "Открываем…" })).toBeDisabled();
    });

    resolveResume(makeResumeDraft());

    await waitFor(() => {
      expect(mockNavigate).toHaveBeenCalledWith("/new?draft=draft-resume-42");
    });
  });

  it("shows error and does not navigate when resume fails", async () => {
    const onClose = vi.fn();
    mockResume.mockRejectedValue(new Error("КП недоступно для дописывания"));
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={onClose} />);

    fireEvent.click(screen.getByRole("button", { name: "Добавить другое наименование" }));

    await waitFor(() => {
      expect(screen.getByText("КП недоступно для дописывания")).toBeInTheDocument();
    });

    expect(mockNavigate).not.toHaveBeenCalled();
    expect(mockDispatch).not.toHaveBeenCalled();
    expect(onClose).not.toHaveBeenCalled();
    expect(screen.getByRole("button", { name: "Добавить другое наименование" })).not.toBeDisabled();
  });
});
