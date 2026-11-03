import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { ContractCardList } from "@/features/commercial-archive/components/ContractCardList";
import type { ContractCard, DirectoryClient } from "@/features/commercial-archive/types/supplyContract";
import { ApiError } from "@/shared/lib/apiError";

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    searchDirectory: vi.fn(),
    ensureDirectoryContract: vi.fn(),
    searchContractCards: vi.fn(),
    createContractCard: vi.fn(),
    getContractCard: vi.fn(),
    patchContractCard: vi.fn(),
    saveContractRequisites: vi.fn(),
    parseContractCard: vi.fn(),
    deleteContractCard: vi.fn(),
  },
}));

const card = (overrides: Partial<ContractCard> = {}): ContractCard => ({
  id: 9,
  name: "УНИСТРОЙ",
  imported_name: "УНИСТРОЙ",
  counterparty_id: 7,
  number: "1032/11/26",
  contract_date: "2026-11-02",
  manager_name: null,
  status: "нет",
  scan_note: null,
  has_scan: false,
  legal_form: null,
  full_name: null,
  short_name: null,
  signatory_position: null,
  signatory_name: null,
  signatory_verb: null,
  authority_basis: null,
  poa_number: null,
  poa_date: null,
  inn: "7701000001",
  kpp: "770101001",
  ogrn: null,
  legal_address: null,
  postal_address: null,
  phone: null,
  email: null,
  bank_name: null,
  account: null,
  corr_account: null,
  bik: null,
  edo_operator: null,
  edo_id: null,
  okved: null,
  requisites_complete: false,
  ...overrides,
});

const clientRow = (overrides: Partial<DirectoryClient> = {}): DirectoryClient => ({
  id: 7,
  name: "УНИСТРОЙ",
  inn: "7701000001",
  kpp: "770101001",
  code: "00-00000077",
  contract: null,
  linkable: null,
  candidates: [],
  ...overrides,
});

const renderList = () => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  return render(
    <QueryClientProvider client={client}>
      <ContractCardList />
    </QueryClientProvider>,
  );
};

const typeSearch = (value: string) => {
  fireEvent.change(screen.getByRole("searchbox", { name: "Поиск" }), { target: { value } });
};

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

beforeEach(() => {
  vi.mocked(archiveApi.searchDirectory).mockResolvedValue({ items: [], count: 0 });
});

describe("ContractCardList", () => {
  it("asks for a name or INN and does not search contract cards", () => {
    renderList();

    expect(screen.getByRole("searchbox", { name: "Поиск" })).toHaveAttribute("placeholder", "Имя или ИНН");
    expect(screen.getByText("Введите имя или ИНН")).toBeInTheDocument();
    expect(screen.queryByRole("form", { name: "Новая карточка" })).not.toBeInTheDocument();
    expect(screen.queryByLabelText("Номер договора")).not.toBeInTheDocument();
    expect(screen.queryByLabelText("Зона файла новой карточки")).not.toBeInTheDocument();
    expect(archiveApi.searchDirectory).not.toHaveBeenCalled();
    expect(archiveApi.searchContractCards).not.toHaveBeenCalled();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });

  it("keeps the hint for one character and searches the directory from two", async () => {
    vi.mocked(archiveApi.searchDirectory).mockResolvedValue({
      items: [clientRow()],
      count: 1,
    });
    renderList();

    typeSearch("у");
    expect(screen.getByText("Введите имя или ИНН")).toBeInTheDocument();
    expect(archiveApi.searchDirectory).not.toHaveBeenCalled();

    typeSearch("уни");

    const list = await screen.findByRole("list", { name: "Контрагенты" });
    expect(archiveApi.searchDirectory).toHaveBeenCalledWith("уни");
    expect(archiveApi.searchContractCards).not.toHaveBeenCalled();
    expect(within(list).getByText("УНИСТРОЙ")).toBeInTheDocument();
    expect(within(list).getByText("7701000001")).toBeInTheDocument();
    expect(within(list).getByText("договор нет")).toBeInTheDocument();
    expect(within(list).getByLabelText("Номер договора")).toBeInTheDocument();
    expect(within(list).queryByRole("button", { name: "Завести" })).not.toBeInTheDocument();
    expect(archiveApi.ensureDirectoryContract).not.toHaveBeenCalled();
  });

  it("shows Привязать with the linkable number", async () => {
    vi.mocked(archiveApi.searchDirectory).mockResolvedValue({
      items: [clientRow({ linkable: { id: 4, number: "0882/09/25" } })],
      count: 1,
    });
    renderList();

    typeSearch("уни");

    expect(await screen.findByRole("button", { name: "Привязать № 0882/09/25" })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Завести" })).not.toBeInTheDocument();
  });

  it("does not send an empty number and shows the typed number after a successful link", async () => {
    const without = clientRow();
    const withNumber = clientRow({ contract: { id: 9, number: "0882/09/25" }, linkable: null, candidates: [] });
    let calls = 0;
    vi.mocked(archiveApi.searchDirectory).mockImplementation(async () => {
      calls += 1;
      return calls === 1 ? { items: [without], count: 1 } : { items: [withNumber], count: 1 };
    });
    vi.mocked(archiveApi.ensureDirectoryContract).mockResolvedValue({ id: 9, number: "0882/09/25" });
    renderList();

    typeSearch("уни");
    const form = await screen.findByRole("form", { name: "Старый номер" });
    fireEvent.submit(form);
    expect(archiveApi.ensureDirectoryContract).not.toHaveBeenCalled();

    fireEvent.change(within(form).getByLabelText("Номер договора"), { target: { value: "  " } });
    fireEvent.submit(form);
    expect(archiveApi.ensureDirectoryContract).not.toHaveBeenCalled();

    fireEvent.change(within(form).getByLabelText("Номер договора"), { target: { value: "0882/09/25" } });
    fireEvent.submit(form);

    await waitFor(() => {
      expect(archiveApi.ensureDirectoryContract).toHaveBeenCalledWith(7, { number: "0882/09/25" });
    });
    expect(archiveApi.ensureDirectoryContract).toHaveBeenCalledTimes(1);
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(archiveApi.searchContractCards).not.toHaveBeenCalled();
    const list = await screen.findByRole("list", { name: "Контрагенты" });
    expect(within(list).getByText("0882/09/25")).toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Завести" })).not.toBeInTheDocument();
  });

  it("sends the chosen candidate id when two cards are listed", async () => {
    vi.mocked(archiveApi.searchDirectory).mockResolvedValue({
      items: [
        clientRow({
          candidates: [
            { id: 4, number: "0100/01/24" },
            { id: 5, number: "0882/09/25" },
          ],
        }),
      ],
      count: 1,
    });
    vi.mocked(archiveApi.ensureDirectoryContract).mockResolvedValue({ id: 5, number: "0882/09/25" });
    renderList();

    typeSearch("уни");
    fireEvent.click(await screen.findByRole("button", { name: "Привязать № 0882/09/25" }));

    await waitFor(() => {
      expect(archiveApi.ensureDirectoryContract).toHaveBeenCalledWith(7, { contract_id: 5 });
    });
    expect(archiveApi.ensureDirectoryContract).not.toHaveBeenCalledWith(7, {});
  });

  it("links from Привязать with the card id and leaves the list open", async () => {
    const without = clientRow({
      linkable: { id: 4, number: "0882/09/25" },
      candidates: [{ id: 4, number: "0882/09/25" }],
    });
    const linked = clientRow({ contract: { id: 4, number: "0882/09/25" }, linkable: null, candidates: [] });
    let calls = 0;
    vi.mocked(archiveApi.searchDirectory).mockImplementation(async () => {
      calls += 1;
      return calls === 1 ? { items: [without], count: 1 } : { items: [linked], count: 1 };
    });
    vi.mocked(archiveApi.ensureDirectoryContract).mockResolvedValue({ id: 4, number: "0882/09/25" });
    renderList();

    typeSearch("7701000001");
    fireEvent.click(await screen.findByRole("button", { name: "Привязать № 0882/09/25" }));

    await waitFor(() => {
      expect(archiveApi.ensureDirectoryContract).toHaveBeenCalledWith(7, { contract_id: 4 });
    });
    expect(archiveApi.searchDirectory).toHaveBeenCalledWith("7701000001");
    const list = await screen.findByRole("list", { name: "Контрагенты" });
    expect(within(list).getByRole("button", { name: /0882\/09\/25/ })).toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });

  it("opens the existing card and keeps the file zone inside it", async () => {
    vi.mocked(archiveApi.searchDirectory).mockResolvedValue({
      items: [clientRow({ contract: { id: 9, number: "0882/09/25" } })],
      count: 1,
    });
    vi.mocked(archiveApi.getContractCard).mockResolvedValue(
      card({ id: 9, number: "0882/09/25", contract_date: "2025-09-08" }),
    );
    renderList();

    typeSearch("уни");
    fireEvent.click(await screen.findByRole("button", { name: /0882\/09\/25/ }));

    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    expect(archiveApi.getContractCard).toHaveBeenCalledWith(9);
    expect(archiveApi.searchContractCards).not.toHaveBeenCalled();
    expect(opened).toHaveTextContent("УНИСТРОЙ");
    expect(opened).toHaveTextContent("0882/09/25");
    expect(within(opened).getByLabelText("Зона файла карточки")).toBeInTheDocument();
    expect(within(opened).getByLabelText("Файл карточки")).toBeInTheDocument();
    expect(screen.getByRole("list", { name: "Контрагенты" })).toBeInTheDocument();
    expect(screen.queryByLabelText("Зона файла новой карточки")).not.toBeInTheDocument();
  });

  it("shows an empty result and a search error", async () => {
    vi.mocked(archiveApi.searchDirectory).mockResolvedValue({ items: [], count: 0 });
    renderList();
    typeSearch("неттакого");
    expect(await screen.findByText("Ничего не найдено")).toBeInTheDocument();

    cleanup();
    vi.mocked(archiveApi.searchDirectory).mockRejectedValue(
      new ApiError("сбой", 500, "Не удалось найти контрагентов"),
    );
    renderList();
    typeSearch("уни");
    expect(await screen.findByRole("alert")).toHaveTextContent("Не удалось найти контрагентов");
  });

  it("opens a blank from Собрать договор and does not create a row", async () => {
    vi.mocked(archiveApi.searchDirectory).mockResolvedValue({
      items: [clientRow()],
      count: 1,
    });
    renderList();

    typeSearch("уни");
    fireEvent.click(await screen.findByRole("button", { name: "Собрать договор" }));

    expect(screen.getByRole("form", { name: "Реквизиты договора" })).toBeInTheDocument();
    expect(screen.queryByRole("form", { name: "Данные карточки" })).not.toBeInTheDocument();
    expect(archiveApi.ensureDirectoryContract).not.toHaveBeenCalled();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(archiveApi.getContractCard).not.toHaveBeenCalled();

    fireEvent.change(screen.getByLabelText("Полное наименование"), { target: { value: "Черновик" } });
    fireEvent.click(screen.getByRole("button", { name: "Закрыть" }));
    expect(screen.queryByRole("form", { name: "Реквизиты договора" })).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Собрать договор" }));
    expect(screen.getByLabelText("Полное наименование")).toHaveValue("");
    expect(archiveApi.ensureDirectoryContract).not.toHaveBeenCalled();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });
});
