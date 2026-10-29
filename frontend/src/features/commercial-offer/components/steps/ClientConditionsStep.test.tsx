import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ClientConditionsStep } from "@/features/commercial-offer/components/steps/ClientConditionsStep";
import type { CounterpartyShort } from "@/features/commercial-offer/api/counterpartiesApi";

const mockUseAuth = vi.fn();
const searchCounterparties = vi.fn();
const searchContractCards = vi.fn();

vi.mock("@/features/auth/model/AuthProvider", () => ({
  useAuth: () => mockUseAuth(),
}));

vi.mock("@/features/commercial-offer/api/counterpartiesApi", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@/features/commercial-offer/api/counterpartiesApi")>();
  return {
    ...actual,
    searchCounterparties: (...args: unknown[]) => searchCounterparties(...args),
    searchContractCards: (...args: unknown[]) => searchContractCards(...args),
  };
});

vi.mock("@/shared/lib/useDebouncedValue", () => ({
  useDebouncedValue: <T,>(value: T) => value,
}));

const managers = [
  { id: 10, fio: "Иванов И.И.", contact_number: "+7 900 000-00-01", email: "ivanov@example.com" },
  { id: 20, fio: "Петров П.П.", contact_number: "+7 900 000-00-02", email: "petrov@example.com" },
];

const ROMA: CounterpartyShort = {
  id: 7,
  code_1c: "00-00000007",
  name: "ООО Ромашка",
  inn: "7701234567",
  kpp: "770101001",
};

const defaultProps = {
  managers,
  selectedManagerId: null as number | null,
  defaultValues: {
    clientName: "",
    counterpartyId: null as number | null,
    conditionsMode: "standard" as const,
    deliveryConditions: "",
    paymentConditions: "",
  },
  errorMessage: null,
  isPending: false,
  onBack: vi.fn(),
  onManagerChange: vi.fn(),
  onSubmit: vi.fn(),
};

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({ defaultOptions: { queries: { retry: false, gcTime: 0 } } });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

describe("ClientConditionsStep", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("searches the 1C directory and submits id, name, inn, kpp and code", async () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "manager", role: "manager", manager_id: 10, is_active: true },
    });
    searchCounterparties.mockResolvedValue({ items: [ROMA], count: 1 });
    const onSubmit = vi.fn();

    wrap(<ClientConditionsStep {...defaultProps} selectedManagerId={10} onSubmit={onSubmit} />);

    fireEvent.change(screen.getByRole("combobox", { name: /клиент/i }), {
      target: { value: "ромаш" },
    });

    expect(searchCounterparties).toHaveBeenCalledWith("ромаш");
    expect(searchContractCards).not.toHaveBeenCalled();
    const option = await screen.findByRole("option", { name: /ООО Ромашка/ });
    expect(option).toHaveTextContent("ИНН 7701234567");
    expect(option).toHaveTextContent("КПП 770101001");
    expect(option).toHaveTextContent("00-00000007");
    expect(option).not.toHaveTextContent("0001/09/26");
    expect(option).not.toHaveTextContent("0006/09/26");

    fireEvent.click(option);

    expect(screen.getByText(/ИНН 7701234567 · КПП 770101001 · код 1С 00-00000007/)).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: /Рассчитать КП/i }));
    await waitFor(() => {
      expect(onSubmit).toHaveBeenCalledWith(
        expect.objectContaining({
          managerId: 10,
          clientName: "ООО Ромашка",
          counterpartyId: 7,
          counterpartyCode1c: "00-00000007",
          counterpartyInn: "7701234567",
          counterpartyKpp: "770101001",
        }),
      );
    });
  });

  it("auto-selects manager from auth profile when manager_id is in list", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "manager", role: "manager", manager_id: 10, is_active: true },
    });

    const onManagerChange = vi.fn();
    wrap(<ClientConditionsStep {...defaultProps} onManagerChange={onManagerChange} />);

    expect(onManagerChange).toHaveBeenCalledWith(10);
    expect(screen.getByText("Иванов И.И.")).toBeInTheDocument();
  });

  it("shows manager select when profile has no manager_id", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "admin", role: "admin", manager_id: null, is_active: true },
    });

    wrap(<ClientConditionsStep {...defaultProps} />);

    expect(screen.getByText("Выберите менеджера для итоговых документов КП.")).toBeInTheDocument();
    expect(screen.getByRole("combobox", { name: /менеджер/i })).toBeInTheDocument();
  });

  it("submits typed name without a directory choice", async () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "manager", role: "manager", manager_id: 10, is_active: true },
    });
    searchCounterparties.mockResolvedValue({ items: [], count: 0 });
    const onSubmit = vi.fn();

    wrap(<ClientConditionsStep {...defaultProps} selectedManagerId={10} onSubmit={onSubmit} />);

    fireEvent.change(screen.getByRole("combobox", { name: /клиент/i }), {
      target: { value: "ИП Петров" },
    });
    expect(searchCounterparties).toHaveBeenCalledWith("ИП Петров");
    expect(searchContractCards).not.toHaveBeenCalled();

    const submit = screen.getByRole("button", { name: /Рассчитать КП/i });
    expect(submit).toBeEnabled();
    fireEvent.click(submit);
    await waitFor(() => {
      expect(onSubmit).toHaveBeenCalledWith(
        expect.objectContaining({
          managerId: 10,
          clientName: "ИП Петров",
          counterpartyId: null,
          counterpartyInn: null,
        }),
      );
    });
  });

  it("shows archive-vs-production hint on the client field", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "manager", role: "manager", manager_id: 10, is_active: true },
    });

    wrap(<ClientConditionsStep {...defaultProps} selectedManagerId={10} />);

    expect(
      screen.getByText(/Без карточки 1С можно сохранить в архив; в производство — только после привязки контрагента/),
    ).toBeInTheDocument();
  });

  it("submits the chosen counterparty and shows requisites card", async () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "manager", role: "manager", manager_id: 10, is_active: true },
    });
    searchCounterparties.mockResolvedValue({ items: [ROMA], count: 1 });
    const onSubmit = vi.fn();

    wrap(<ClientConditionsStep {...defaultProps} selectedManagerId={10} onSubmit={onSubmit} />);

    fireEvent.change(screen.getByRole("combobox", { name: /клиент/i }), {
      target: { value: "ромаш" },
    });
    fireEvent.click(await screen.findByRole("option", { name: /ИНН 7701234567/ }));

    expect(screen.getByText(/ИНН 7701234567 · КПП 770101001 · код 1С 00-00000007/)).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: /Рассчитать КП/i }));
    await waitFor(() => {
      expect(onSubmit).toHaveBeenCalledWith(
        expect.objectContaining({
          managerId: 10,
          clientName: "ООО Ромашка",
          counterpartyId: 7,
        }),
      );
    });
  });

  it("hydrates the selected card from draft counterpartyId", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "manager", role: "manager", manager_id: 10, is_active: true },
    });

    wrap(
      <ClientConditionsStep
        {...defaultProps}
        selectedManagerId={10}
        defaultValues={{
          ...defaultProps.defaultValues,
          clientName: "ООО Ромашка",
          counterpartyId: 7,
          counterpartyCode1c: "00-00000007",
          counterpartyInn: "7701234567",
          counterpartyKpp: "770101001",
        }}
      />,
    );

    expect(screen.getByRole("combobox", { name: /клиент/i })).toHaveValue("ООО Ромашка");
    expect(screen.getByText(/ИНН 7701234567 · КПП 770101001 · код 1С 00-00000007/)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Рассчитать КП/i })).toBeEnabled();
  });
});
