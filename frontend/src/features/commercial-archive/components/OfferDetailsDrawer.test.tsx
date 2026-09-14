import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { OfferDetailsDrawer } from "@/features/commercial-archive/components/OfferDetailsDrawer";
import type { ArchiveOfferDetails, KpReadinessSummary } from "@/features/commercial-archive/types/archive";

const mockUseArchiveOfferQuery = vi.fn();
const mockUsePromiseHoldQuery = vi.fn(() => ({ data: null, isPending: false, isError: false }));
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

vi.mock("@/features/factory-capacity/api/promiseQuote", async () => {
  const actual = await vi.importActual<typeof import("@/features/factory-capacity/api/promiseQuote")>(
    "@/features/factory-capacity/api/promiseQuote",
  );
  return {
    ...actual,
    usePromiseHoldQuery: (...args: unknown[]) => mockUsePromiseHoldQuery(...args),
  };
});

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    resume: (...args: unknown[]) => mockResume(...args),
    buildDocumentUrl: (kpId: number, kind: string) =>
      `/api/v1/commercial/archive/${kpId}/files/${kind}`,
  },
}));

const mockDownloadFile = vi.fn();
vi.mock("@/shared/lib/downloadFile", () => ({
  downloadFile: (...args: unknown[]) => mockDownloadFile(...args),
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
    saved_offer: { kp_id: 42, status: "в архиве", execution_terms: "" },
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

  it("shows editable delivery schedule button for в работе", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе", makeReadiness()),
      isPending: false,
      isError: false,
      error: null,
    });
    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);
    expect(screen.getByRole("button", { name: "График поставки" })).toBeEnabled();
    expect(screen.getByRole("button", { name: "График поставки" })).toHaveAttribute(
      "title",
      "Редактирование графика поставки",
    );
  });

  it("shows read-only delivery schedule button for На СГП and выполнено", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("На СГП", makeReadiness()),
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
      data: makeOffer("выполнено", makeReadiness()),
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

  it("hides delivery schedule button when status is в архиве", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", makeReadiness()),
      isPending: false,
      isError: false,
      error: null,
    });
    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);
    expect(screen.queryByRole("button", { name: "График поставки" })).not.toBeInTheDocument();
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

describe("OfferDetailsDrawer archive constructor CTAs", () => {
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

  it("renders dual constructor CTAs when status is в архиве", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByRole("button", { name: "(+ Добавить)" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Редактировать" })).toBeInTheDocument();
  });

  it.each(["в работе", "выполнено", "На СГП", "отклонено", "в ожидании"])(
    "does not render constructor CTAs when status is %s",
    (status) => {
      mockUseArchiveOfferQuery.mockReturnValue({
        data: makeOffer(status),
        isPending: false,
        isError: false,
        error: null,
      });

      render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

      expect(screen.queryByRole("button", { name: "(+ Добавить)" })).not.toBeInTheDocument();
      expect(screen.queryByRole("button", { name: "Редактировать" })).not.toBeInTheDocument();
      expect(
        screen.queryByRole("button", { name: "Добавить другое наименование" }),
      ).not.toBeInTheDocument();
    },
  );

  it("(+ Добавить) resumes draft, hydrates, starts append cycle, and navigates", async () => {
    const onClose = vi.fn();
    const draft = makeResumeDraft("draft-resume-42");
    mockResume.mockResolvedValue(draft);
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={onClose} />);

    fireEvent.click(screen.getByRole("button", { name: "(+ Добавить)" }));

    await waitFor(() => {
      expect(mockResume).toHaveBeenCalledWith(42);
    });

    await waitFor(() => {
      expect(mockDispatch).toHaveBeenCalledWith({ type: "hydrate-draft", payload: draft });
      expect(mockDispatch).toHaveBeenCalledWith({ type: "start-append-cycle" });
      expect(mockNavigate).toHaveBeenCalledWith("/new?draft=draft-resume-42");
      expect(onClose).toHaveBeenCalled();
    });

    expect(mockDispatch).not.toHaveBeenCalledWith({ type: "set-step", step: "result" });
    expect(mockDispatch).not.toHaveBeenCalledWith({ type: "set-save-result", payload: null });

    const hydrateIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "hydrate-draft",
    );
    const appendIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "start-append-cycle",
    );
    expect(hydrateIndex).toBeGreaterThanOrEqual(0);
    expect(appendIndex).toBeGreaterThan(hydrateIndex);
  });

  it("Редактировать resumes draft, hydrates to result, without append cycle", async () => {
    const onClose = vi.fn();
    const draft = makeResumeDraft("draft-edit-42");
    mockResume.mockResolvedValue(draft);
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={onClose} />);

    fireEvent.click(screen.getByRole("button", { name: "Редактировать" }));

    await waitFor(() => {
      expect(mockResume).toHaveBeenCalledWith(42);
    });

    await waitFor(() => {
      expect(mockDispatch).toHaveBeenCalledWith({ type: "hydrate-draft", payload: draft });
      expect(mockDispatch).toHaveBeenCalledWith({ type: "set-step", step: "result" });
      expect(mockNavigate).toHaveBeenCalledWith("/new?draft=draft-edit-42");
      expect(onClose).toHaveBeenCalled();
    });

    expect(mockDispatch).not.toHaveBeenCalledWith({ type: "start-append-cycle" });

    const hydrateIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "hydrate-draft",
    );
    const setStepIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "set-step" && call[0]?.step === "result",
    );
    expect(hydrateIndex).toBeGreaterThanOrEqual(0);
    expect(setStepIndex).toBeGreaterThan(hydrateIndex);
  });

  it("Редактировать сбрасывает lastSaveResult до hydrate-draft и set-step result", async () => {
    const draft = makeResumeDraft("draft-edit-save-flag");
    mockResume.mockResolvedValue(draft);
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    fireEvent.click(screen.getByRole("button", { name: "Редактировать" }));

    await waitFor(() => {
      expect(mockDispatch).toHaveBeenCalledWith({ type: "set-save-result", payload: null });
      expect(mockDispatch).toHaveBeenCalledWith({ type: "hydrate-draft", payload: draft });
      expect(mockDispatch).toHaveBeenCalledWith({ type: "set-step", step: "result" });
    });

    const resetIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "set-save-result" && call[0]?.payload === null,
    );
    const hydrateIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "hydrate-draft",
    );
    const setStepIndex = mockDispatch.mock.calls.findIndex(
      (call) => call[0]?.type === "set-step" && call[0]?.step === "result",
    );
    expect(resetIndex).toBeGreaterThanOrEqual(0);
    expect(hydrateIndex).toBeGreaterThan(resetIndex);
    expect(setStepIndex).toBeGreaterThan(hydrateIndex);
  });

  it("disables both CTAs while resume is pending", async () => {
    let resolveResume: (value: unknown) => void = () => undefined;
    mockResume.mockReturnValue(
      new Promise((resolve) => {
        resolveResume = resolve;
      }),
    );
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    fireEvent.click(screen.getByRole("button", { name: "(+ Добавить)" }));

    await waitFor(() => {
      const pendingButtons = screen.getAllByRole("button", { name: "Открываем…" });
      expect(pendingButtons.length).toBeGreaterThanOrEqual(1);
      pendingButtons.forEach((btn) => expect(btn).toBeDisabled());
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
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={onClose} />);

    fireEvent.click(screen.getByRole("button", { name: "Редактировать" }));

    await waitFor(() => {
      expect(screen.getByText("КП недоступно для дописывания")).toBeInTheDocument();
    });

    expect(mockNavigate).not.toHaveBeenCalled();
    expect(mockDispatch).not.toHaveBeenCalled();
    expect(onClose).not.toHaveBeenCalled();
    expect(screen.getByRole("button", { name: "(+ Добавить)" })).not.toBeDisabled();
    expect(screen.getByRole("button", { name: "Редактировать" })).not.toBeDisabled();
  });

  it("shows read-only Итоги without finance inputs for в архиве", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", null, {
        logistics_cost: 15000,
        total_cargo_weight_kg: 1200,
        delivery_service_total_rub: 15000,
        finance: {
          subtotal: 1000,
          vat_amount: 220,
          total_amount: 1220,
          discount_percent: 5,
        },
      }),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByText("Итоги")).toBeInTheDocument();
    expect(screen.getByText("Общий вес груза, кг")).toBeInTheDocument();
    expect(screen.getByText("НДС (22%)")).toBeInTheDocument();
    expect(screen.getByText(/Услуга по доставке грузов/)).toBeInTheDocument();
    expect(screen.getByText("Скидка")).toBeInTheDocument();
    expect(screen.getByText("Итого с НДС")).toBeInTheDocument();

    expect(screen.queryByPlaceholderText("Стоимость одного рейса")).not.toBeInTheDocument();
    expect(screen.queryByPlaceholderText("Например, 2 000 000")).not.toBeInTheDocument();
    expect(screen.queryByPlaceholderText("Например, 5")).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "OK" })).not.toBeInTheDocument();
  });
});

describe("OfferDetailsDrawer promise hold badge", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
    mockUsePromiseHoldQuery.mockReturnValue({ data: null, isPending: false, isError: false });
  });

  it("shows hold badge with who pinned the date", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });
    mockUsePromiseHoldQuery.mockReturnValue({
      data: {
        id: 1,
        kp_id: 42,
        kind: "hold",
        status: "active",
        tracks_total: 2,
        promised_date: "2026-09-18",
        expires_at: "2026-09-03T23:59:59",
        created_by: "alice",
        created_at: "2026-09-03T12:00:00",
        allocations: [],
      },
      isPending: false,
      isError: false,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    const badge = screen.getByTestId("promise-hold-badge");
    expect(badge).toHaveTextContent("к 18.09 · до вечера");
    expect(badge).toHaveAttribute("title", "Закрепил: alice");
  });

  it("hides hold badge when there is no active hold", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);
    expect(screen.queryByTestId("promise-hold-badge")).not.toBeInTheDocument();
  });
});

describe("OfferDetailsDrawer counterparty requisites", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("does not show INN/KPP when the offer has no requisites", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве"),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByText("ООО Тест")).toBeInTheDocument();
    expect(screen.queryByText(/ИНН /)).not.toBeInTheDocument();
  });

  it("shows INN/KPP next to the client when they are present", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", null, {
        customer_inn: "7701234567",
        customer_kpp: "770101001",
      }),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByText("ИНН 7701234567 / КПП 770101001")).toBeInTheDocument();
  });
});

describe("OfferDetailsDrawer embed-delivery XLSX button", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("shows the embed button next to ordinary XLSX and disables it when delivery is 0", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", null, { delivery_service_total_rub: 0 }),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    const embedButton = screen.getByRole("button", { name: "XLSX (доставка в цене)" });
    const plainXlsx = screen.getByRole("button", { name: /XLSX$/ });
    expect(screen.getByRole("button", { name: /PDF/ })).toBeEnabled();
    expect(plainXlsx).toBeEnabled();
    expect(embedButton).toBeDisabled();
    expect(embedButton).toHaveAttribute("title", "Нет суммы доставки");
  });

  it("downloads xlsx_delivery_in_unit when delivery is positive", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", null, { delivery_service_total_rub: 156800 }),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    const embedButton = screen.getByRole("button", { name: "XLSX (доставка в цене)" });
    expect(embedButton).toBeEnabled();
    fireEvent.click(embedButton);
    expect(mockDownloadFile).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/42/files/xlsx_delivery_in_unit",
    );
  });
});

describe("OfferDetailsDrawer FBS/LM delivery boiler", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("shows trips and pending marks for a new FBS offer with the flag on", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", null, {
        product_type: "fbs",
        fbs_lm_delivery_enabled: true,
        fbs_lm_delivery_ready: true,
        fbs_lm_trips: 2,
        fbs_lm_delivery_total: 100000,
        delivery_service_total_rub: 100000,
        fbs: [
          {
            position_number: 1,
            mark: "ФБС 24.4.6-Т",
            concrete_grade: "",
            qty: 10,
            unit_price: 1200,
            discounted_price: 1200,
          },
        ],
      }),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByText("Рейсов ФБС/ЛС/ЛМ")).toBeInTheDocument();
    expect(screen.getByText("2")).toBeInTheDocument();
    expect(screen.queryByText(/Нет веса в справочнике/)).not.toBeInTheDocument();
    const embedButton = screen.getByRole("button", { name: "XLSX (доставка в цене)" });
    expect(embedButton).toBeEnabled();
  });

  it("shows pending marks and hides trips when the boiler is not ready", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в работе", null, {
        product_type: "fbs",
        fbs_lm_delivery_enabled: true,
        fbs_lm_delivery_ready: false,
        fbs_lm_trips: 0,
        fbs_lm_pending_marks: ["ЛС99"],
        fbs: [
          {
            position_number: 1,
            mark: "ЛС99",
            concrete_grade: "",
            qty: 4,
            unit_price: 800,
            discounted_price: 800,
          },
        ],
      }),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.getByText("Нет веса в справочнике — доставка ФБС/ЛС/ЛМ не посчитана")).toBeInTheDocument();
    expect(screen.getAllByText("ЛС99").length).toBeGreaterThan(1);
    expect(screen.queryByText("Рейсов ФБС/ЛС/ЛМ")).not.toBeInTheDocument();
  });

  it("does not show the FBS/LM boiler on an old offer without the flag", () => {
    mockUseArchiveOfferQuery.mockReturnValue({
      data: makeOffer("в архиве", null, {
        product_type: "fbs",
        fbs_lm_delivery_enabled: false,
        fbs_lm_delivery_ready: true,
        fbs_lm_trips: 0,
        fbs_lm_pending_marks: [],
        delivery_service_total_rub: 0,
        fbs: [
          {
            position_number: 1,
            mark: "ФБС 24.4.6-Т",
            concrete_grade: "",
            qty: 10,
            unit_price: 1200,
            discounted_price: 1200,
          },
        ],
      }),
      isPending: false,
      isError: false,
      error: null,
    });

    render(<OfferDetailsDrawer open kpId={42} onClose={vi.fn()} />);

    expect(screen.queryByText("Рейсов ФБС/ЛС/ЛМ")).not.toBeInTheDocument();
    expect(screen.queryByText(/Нет веса в справочнике/)).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "XLSX (доставка в цене)" })).toBeDisabled();
  });
});
