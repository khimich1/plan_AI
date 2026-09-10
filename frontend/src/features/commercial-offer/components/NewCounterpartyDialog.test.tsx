import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { NewCounterpartyDialog } from "@/features/commercial-offer/components/NewCounterpartyDialog";
import {
  createCounterparty,
  DuplicateCounterpartyError,
  type CounterpartyShort,
} from "@/features/commercial-offer/api/counterpartiesApi";

vi.mock("@/features/commercial-offer/api/counterpartiesApi", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@/features/commercial-offer/api/counterpartiesApi")>();
  return {
    ...actual,
    createCounterparty: vi.fn(),
  };
});

const created: CounterpartyShort = {
  id: 88,
  code_1c: "00-00000088",
  name: "Новый клиент",
  inn: "7701999888",
  kpp: "770101001",
};

const existing: CounterpartyShort = {
  id: 12,
  code_1c: "00-00000012",
  name: "Уже есть ООО",
  inn: "7701111222",
  kpp: "770101002",
};

describe("NewCounterpartyDialog", () => {
  beforeEach(() => {
    vi.mocked(createCounterparty).mockReset();
  });

  afterEach(() => {
    cleanup();
  });

  it("asks to copy the 1C code and requires name and code", () => {
    render(<NewCounterpartyDialog open onClose={vi.fn()} onCreated={vi.fn()} />);

    expect(
      screen.getByText(/Сначала заведите контрагента в 1С и скопируйте код/),
    ).toBeInTheDocument();
    expect(screen.getByLabelText("Наименование")).toBeInTheDocument();
    expect(screen.getByLabelText("Код 1С")).toBeInTheDocument();
    expect(screen.getByLabelText("ИНН")).toBeInTheDocument();
    expect(screen.getByLabelText("КПП")).toBeInTheDocument();
    expect(screen.getByLabelText("Клиент")).toBeChecked();

    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));
    expect(createCounterparty).not.toHaveBeenCalled();
  });

  it("creates a counterparty and selects it", async () => {
    vi.mocked(createCounterparty).mockResolvedValue({ item: created });
    const onCreated = vi.fn();
    const onClose = vi.fn();
    render(<NewCounterpartyDialog open onClose={onClose} onCreated={onCreated} initialName="Новый клиент" />);

    fireEvent.change(screen.getByLabelText("Код 1С"), { target: { value: "00-00000088" } });
    fireEvent.change(screen.getByLabelText("ИНН"), { target: { value: "7701999888" } });
    fireEvent.change(screen.getByLabelText("КПП"), { target: { value: "770101001" } });
    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));

    await waitFor(() => {
      expect(createCounterparty).toHaveBeenCalledWith({
        name: "Новый клиент",
        code_1c: "00-00000088",
        inn: "7701999888",
        kpp: "770101001",
        is_client: true,
      });
      expect(onCreated).toHaveBeenCalledWith(created);
      expect(onClose).toHaveBeenCalled();
    });
  });

  it("shows an inn-duplicate warning without blocking selection", async () => {
    vi.mocked(createCounterparty).mockResolvedValue({
      item: created,
      warning: "Контрагент с таким ИНН уже есть",
    });
    const onCreated = vi.fn();
    const onClose = vi.fn();
    render(<NewCounterpartyDialog open onClose={onClose} onCreated={onCreated} initialName="Новый клиент" />);

    fireEvent.change(screen.getByLabelText("Код 1С"), { target: { value: "00-00000088" } });
    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));

    expect(await screen.findByText(/Контрагент с таким ИНН уже есть/)).toBeInTheDocument();
    expect(onCreated).toHaveBeenCalledWith(created);
    expect(onClose).not.toHaveBeenCalled();
  });

  it("shows the existing record on 409 and lets the user select it", async () => {
    vi.mocked(createCounterparty).mockRejectedValue(
      new DuplicateCounterpartyError("Контрагент с таким кодом 1С уже есть", existing),
    );
    const onCreated = vi.fn();
    const onClose = vi.fn();
    render(<NewCounterpartyDialog open onClose={onClose} onCreated={onCreated} initialName="Уже есть ООО" />);

    fireEvent.change(screen.getByLabelText("Код 1С"), { target: { value: "00-00000012" } });
    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));

    expect(await screen.findByText(/Уже есть ООО/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Выбрать её" }));
    expect(onCreated).toHaveBeenCalledWith(existing);
    expect(onClose).toHaveBeenCalled();
  });
});
