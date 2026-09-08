import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CounterpartyAutocomplete } from "@/features/commercial-offer/components/CounterpartyAutocomplete";
import type { CounterpartyShort } from "@/features/commercial-offer/api/counterpartiesApi";

const searchCounterparties = vi.fn();
const createCounterparty = vi.fn();

vi.mock("@/features/commercial-offer/api/counterpartiesApi", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@/features/commercial-offer/api/counterpartiesApi")>();
  return {
    ...actual,
    searchCounterparties: (...args: unknown[]) => searchCounterparties(...args),
    createCounterparty: (...args: unknown[]) => createCounterparty(...args),
  };
});

vi.mock("@/shared/lib/useDebouncedValue", () => ({
  useDebouncedValue: <T,>(value: T) => value,
}));

const ROMA_A: CounterpartyShort = {
  id: 1,
  code_1c: "00-00000001",
  name: "Ромашка",
  inn: "7701234567",
  kpp: "770101001",
};

const ROMA_B: CounterpartyShort = {
  id: 2,
  code_1c: "00-00000002",
  name: "Ромашка",
  inn: "5401112233",
  kpp: "540101001",
};

const STOLICA: CounterpartyShort = {
  id: 3,
  code_1c: "00-00000003",
  name: "СТОЛИЦА ООО",
  inn: "7700000001",
  kpp: "770001001",
};

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false, gcTime: 0 } },
  });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

describe("CounterpartyAutocomplete", () => {
  beforeEach(() => {
    searchCounterparties.mockReset();
    createCounterparty.mockReset();
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("does not search until two characters are typed", async () => {
    wrap(<CounterpartyAutocomplete selected={null} onSelect={vi.fn()} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "Р" } });

    await waitFor(() => {
      expect(searchCounterparties).not.toHaveBeenCalled();
    });
  });

  it("shows name duplicates distinguished by INN", async () => {
    searchCounterparties.mockResolvedValue({ items: [ROMA_A, ROMA_B], count: 2 });
    wrap(<CounterpartyAutocomplete selected={null} onSelect={vi.fn()} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "Ромашка" } });

    expect(await screen.findByRole("option", { name: /ИНН 7701234567/ })).toBeInTheDocument();
    expect(screen.getByRole("option", { name: /ИНН 5401112233/ })).toBeInTheDocument();
    expect(screen.getByRole("option", { name: /КПП 770101001/ })).toBeInTheDocument();
  });

  it("selects an option on click", async () => {
    searchCounterparties.mockResolvedValue({ items: [ROMA_A, ROMA_B], count: 2 });
    const onSelect = vi.fn();
    wrap(<CounterpartyAutocomplete selected={null} onSelect={onSelect} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "Ромашка" } });
    fireEvent.click(await screen.findByRole("option", { name: /ИНН 5401112233/ }));

    expect(onSelect).toHaveBeenCalledWith(ROMA_B);
  });

  it("clears selection when the typed text is edited after a choice", async () => {
    const onSelect = vi.fn();
    wrap(<CounterpartyAutocomplete selected={ROMA_A} onSelect={onSelect} />);

    const input = screen.getByRole("combobox");
    expect(input).toHaveValue("Ромашка");

    fireEvent.change(input, { target: { value: "Ромашка!" } });

    expect(onSelect).toHaveBeenCalledWith(null);
  });

  it("auto-selects a unique exact name match without a click", async () => {
    searchCounterparties.mockResolvedValue({ items: [STOLICA], count: 1 });
    const onSelect = vi.fn();
    wrap(<CounterpartyAutocomplete selected={null} onSelect={onSelect} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: " столица ооо " } });

    await waitFor(() => {
      expect(onSelect).toHaveBeenCalledWith(STOLICA);
    });
  });

  it("does not auto-select when several clients share the same name", async () => {
    searchCounterparties.mockResolvedValue({ items: [ROMA_A, ROMA_B], count: 2 });
    const onSelect = vi.fn();
    wrap(<CounterpartyAutocomplete selected={null} onSelect={onSelect} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "Ромашка" } });
    await screen.findByRole("option", { name: /ИНН 7701234567/ });

    expect(onSelect).not.toHaveBeenCalled();
  });

  it("shows a loading state while search is in flight", async () => {
    searchCounterparties.mockReturnValue(new Promise(() => undefined));
    wrap(<CounterpartyAutocomplete selected={null} onSelect={vi.fn()} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "сто" } });

    expect(await screen.findByText("Ищем…")).toBeInTheDocument();
  });

  it("shows add-counterparty action when search is empty", async () => {
    searchCounterparties.mockResolvedValue({ items: [], count: 0 });
    wrap(<CounterpartyAutocomplete selected={null} onSelect={vi.fn()} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "неттаких" } });

    expect(await screen.findByText("Не найдено среди клиентов")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Добавить контрагента/i })).toBeInTheDocument();
  });

  it("opens the new-counterparty dialog from the empty state", async () => {
    searchCounterparties.mockResolvedValue({ items: [], count: 0 });
    wrap(<CounterpartyAutocomplete selected={null} onSelect={vi.fn()} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "неттаких" } });
    fireEvent.click(await screen.findByRole("button", { name: /Добавить контрагента/i }));

    expect(screen.getByRole("dialog")).toBeInTheDocument();
    expect(screen.getByText(/Сначала заведите контрагента в 1С и скопируйте код/)).toBeInTheDocument();
  });

  it("keeps the dialog open on an INN-duplicate warning so the banner is visible", async () => {
    searchCounterparties.mockResolvedValue({ items: [], count: 0 });
    createCounterparty.mockResolvedValue({
      item: ROMA_A,
      warning: "Контрагент с таким ИНН уже есть",
    });
    const onSelect = vi.fn();
    wrap(<CounterpartyAutocomplete selected={null} onSelect={onSelect} />);

    fireEvent.change(screen.getByRole("combobox"), { target: { value: "неттаких" } });
    fireEvent.click(await screen.findByRole("button", { name: /Добавить контрагента/i }));
    fireEvent.change(screen.getByLabelText("Код 1С"), { target: { value: "00-00000001" } });
    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));

    expect(await screen.findByText(/Контрагент с таким ИНН уже есть/)).toBeInTheDocument();
    expect(screen.getByRole("dialog")).toBeInTheDocument();
    expect(onSelect).toHaveBeenCalledWith(ROMA_A);
  });

  it("selects the highlighted option with Enter and closes on Escape", async () => {
    searchCounterparties.mockResolvedValue({ items: [ROMA_A, ROMA_B], count: 2 });
    const onSelect = vi.fn();
    wrap(<CounterpartyAutocomplete selected={null} onSelect={onSelect} />);

    const input = screen.getByRole("combobox");
    fireEvent.change(input, { target: { value: "Ромашка" } });
    await screen.findByRole("option", { name: /ИНН 7701234567/ });

    fireEvent.keyDown(input, { key: "Enter" });
    expect(onSelect).toHaveBeenCalledWith(ROMA_A);

    fireEvent.keyDown(input, { key: "Escape" });
    expect(screen.queryByRole("listbox")).not.toBeInTheDocument();
  });
});
