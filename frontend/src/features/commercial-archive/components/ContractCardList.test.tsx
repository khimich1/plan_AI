import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { ContractCardList } from "@/features/commercial-archive/components/ContractCardList";
import type { ContractCard } from "@/features/commercial-archive/types/supplyContract";
import { ApiError } from "@/shared/lib/apiError";

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    searchContractCards: vi.fn(),
    findDuplicateContractCards: vi.fn(),
    getContractCard: vi.fn(),
    createContractCard: vi.fn(),
    createSupplyContract: vi.fn(),
    patchContractCard: vi.fn(),
    patchSupplyContract: vi.fn(),
    saveContractRequisites: vi.fn(),
    parseContractCard: vi.fn(),
    parseContractCardFile: vi.fn(),
    downloadContractDocument: vi.fn(),
    downloadContractEdoAgreement: vi.fn(),
    deleteContractCard: vi.fn(),
  },
}));

const card = (overrides: Partial<ContractCard> = {}): ContractCard => ({
  id: 1,
  name: "Секлюцкий",
  imported_name: "Секлюцкий",
  counterparty_id: null,
  number: "0793/02/25",
  contract_date: "2025-02-01",
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
  inn: "440117229007",
  kpp: null,
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

const SEKLYUTSKY = [
  card({ id: 1, number: "0793/02/25" }),
  card({ id: 2, number: "0794/02/25" }),
  card({ id: 3, number: "0933/03/26" }),
];

const renderList = () => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  const view = render(
    <QueryClientProvider client={client}>
      <ContractCardList />
    </QueryClientProvider>,
  );
  return view;
};

const typeSearch = (value: string) => {
  fireEvent.change(screen.getByRole("searchbox", { name: "Поиск" }), { target: { value } });
};

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

beforeEach(() => {
  vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: [], count: 0 });
  vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({ items: [], count: 0 });
});

describe("ContractCardList", () => {
  it("shows не полные on an incomplete card and hides it when requisites are complete", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({
      items: [
        card({ id: 1, name: "Импорт", number: "0793/02/25", requisites_complete: false }),
        card({ id: 2, name: "Ромашка", number: "0001/10/26", requisites_complete: true }),
      ],
      count: 2,
    });
    renderList();

    const list = await screen.findByRole("list", { name: "Карточки договоров" });
    const incomplete = within(list).getByRole("button", { name: /0793\/02\/25/ });
    const complete = within(list).getByRole("button", { name: /0001\/10\/26/ });

    expect(within(incomplete).getByText("не полные")).toBeInTheDocument();
    expect(within(complete).queryByText("не полные")).not.toBeInTheDocument();
  });

  it("lists every saved card under the new-card form", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: SEKLYUTSKY, count: 3 });
    renderList();

    const list = await screen.findByRole("list", { name: "Карточки договоров" });
    expect(archiveApi.searchContractCards).toHaveBeenCalledWith("");
    expect(within(list).getAllByRole("button")).toHaveLength(3);
    const form = screen.getByRole("form", { name: "Новая карточка" });
    expect(form.compareDocumentPosition(list) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
  });

  it("keeps the full list when the query is one character", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: SEKLYUTSKY, count: 3 });
    renderList();

    expect(await screen.findByText("0793/02/25")).toBeInTheDocument();
    typeSearch("с");

    expect(screen.getByText("0793/02/25")).toBeInTheDocument();
    expect(archiveApi.searchContractCards).not.toHaveBeenCalledWith("с");
  });

  it("searches for секлюц and shows three cards with name, INN and number", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: SEKLYUTSKY, count: 3 });
    renderList();

    typeSearch("секлюц");

    const list = await screen.findByRole("list", { name: "Карточки договоров" });
    expect(archiveApi.searchContractCards).toHaveBeenCalledWith("секлюц");
    expect(within(list).getAllByRole("button")).toHaveLength(3);
    expect(within(list).getByRole("button", { name: /0793\/02\/25/ })).toHaveTextContent("Секлюцкий");
    expect(within(list).getByRole("button", { name: /0793\/02\/25/ })).toHaveTextContent("440117229007");
    expect(within(list).getByText("0794/02/25")).toBeInTheDocument();
    expect(within(list).getByText("0933/03/26")).toBeInTheDocument();
  });

  it("shows an empty result instead of the full list when nothing matches", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: [], count: 0 });
    renderList();

    typeSearch("неттакого");

    expect(await screen.findByText("Ничего не найдено")).toBeInTheDocument();
    expect(screen.queryByText("Карточек пока нет")).not.toBeInTheDocument();
  });

  it("shows the search error in an alert", async () => {
    vi.mocked(archiveApi.searchContractCards).mockRejectedValue(
      new ApiError("сбой", 500, "Не удалось найти карточки"),
    );
    renderList();

    typeSearch("секлюц");

    expect(await screen.findByRole("alert")).toHaveTextContent("Не удалось найти карточки");
  });

  it("shows a loading state while the search is in flight", async () => {
    vi.mocked(archiveApi.searchContractCards).mockReturnValue(new Promise(() => undefined));
    renderList();

    typeSearch("секлюц");

    expect(await screen.findByText("Ищу карточки…")).toBeInTheDocument();
  });

  it("opens the card from the list without loading the contract drawer", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: SEKLYUTSKY, count: 3 });
    renderList();

    typeSearch("секлюц");
    fireEvent.click(await screen.findByRole("button", { name: /0793\/02\/25/ }));

    const opened = screen.getByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("Секлюцкий");
    expect(opened).toHaveTextContent("0793/02/25");
    expect(opened).toHaveTextContent("440117229007");
    expect(screen.getByRole("button", { name: /0793\/02\/25/ })).toHaveAttribute("aria-pressed", "true");
    expect(archiveApi.getContractCard).not.toHaveBeenCalled();
  });

  it("shows the contract block on the selected card without creating a KP contract", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: SEKLYUTSKY, count: 3 });
    renderList();

    typeSearch("секлюц");
    fireEvent.click(await screen.findByRole("button", { name: /0793\/02\/25/ }));

    const block = screen.getByRole("region", { name: "Блок договора" });
    expect(block).toHaveTextContent("Дата: 01.02.2025");
    expect(screen.getByLabelText("Имя контрагента")).toHaveValue("Секлюцкий");
    expect(screen.getByLabelText("ИНН контрагента")).toHaveValue("440117229007");
    expect(screen.getByLabelText("Номер")).toHaveValue("0793/02/25");
    expect(screen.getByRole("button", { name: "Сохранить реквизиты" })).toBeDisabled();
    expect(archiveApi.getContractCard).not.toHaveBeenCalled();
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();
  });

  it("updates the selected card after a number change and keeps the server date", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: SEKLYUTSKY, count: 3 });
    vi.mocked(archiveApi.patchContractCard).mockResolvedValue(
      card({ id: 1, name: "Новое имя", number: "1035/10/26", contract_date: "2025-02-01" }),
    );
    renderList();

    typeSearch("секлюц");
    fireEvent.click(await screen.findByRole("button", { name: /0793\/02\/25/ }));
    fireEvent.change(screen.getByLabelText("Номер"), { target: { value: "1035/10/26" } });
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Имя контрагента"), { target: { value: "Новое имя" } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить карточку" }));

    await waitFor(() => {
      expect(archiveApi.patchContractCard).toHaveBeenCalledWith(1, {
        name: "Новое имя",
        inn: "440117229007",
        number: "1035/10/26",
      });
    });
    expect(await screen.findByRole("heading", { name: "Новое имя" })).toBeInTheDocument();
    const opened = screen.getByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("1035/10/26");
    expect(opened).toHaveTextContent("Дата: 01.02.2025");
    expect(opened).not.toHaveTextContent("01.10.2026");
    expect(archiveApi.patchContractCard.mock.calls[0]?.[1]).not.toHaveProperty("contract_date");
  });

  it("creates a card with only a name and opens the returned card", async () => {
    const created = card({
      id: 9,
      name: "Секлюц",
      imported_name: "Секлюц",
      number: "1034/10/26",
      inn: null,
      contract_date: "2026-10-20",
    });
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(created);
    renderList();

    expect(screen.getByRole("button", { name: "Создать карточку" })).toBeDisabled();
    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "  Секлюц  " } });
    fireEvent.change(screen.getByLabelText("ИНН"), { target: { value: "   " } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));

    await waitFor(() => {
      expect(archiveApi.createContractCard).toHaveBeenCalledWith({ name: "Секлюц" });
    });
    expect(archiveApi.findDuplicateContractCards).toHaveBeenCalledWith("Секлюц", "");
    expect(archiveApi.findDuplicateContractCards.mock.invocationCallOrder[0]).toBeLessThan(
      archiveApi.createContractCard.mock.invocationCallOrder[0],
    );
    expect(archiveApi.createContractCard).toHaveBeenCalledTimes(1);
    expect(archiveApi.searchContractCards).toHaveBeenCalledWith("");

    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("Секлюц");
    expect(opened).toHaveTextContent("1034/10/26");
    expect(opened).toHaveTextContent("ИНН не указан");
  });

  it("sends an optional INN and a number the manager typed before save", async () => {
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(
      card({ id: 10, name: "Новый", number: "0786/1/25", inn: "440117229007" }),
    );
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "Новый" } });
    fireEvent.change(screen.getByLabelText("ИНН"), { target: { value: " 440117229007 " } });
    fireEvent.change(screen.getByLabelText("Номер договора"), { target: { value: " 0786/1/25 " } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));

    await waitFor(() => {
      expect(archiveApi.createContractCard).toHaveBeenCalledWith({
        name: "Новый",
        inn: "440117229007",
        number: "0786/1/25",
      });
    });
    expect(archiveApi.findDuplicateContractCards).toHaveBeenCalledWith("Новый", "440117229007");
  });

  it("does not create a card when the name is blank", () => {
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "   " } });
    expect(screen.getByRole("button", { name: "Создать карточку" })).toBeDisabled();
    fireEvent.submit(screen.getByRole("form", { name: "Новая карточка" }));

    expect(archiveApi.findDuplicateContractCards).not.toHaveBeenCalled();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });

  it("shows a duplicate number error in an alert and does not open a card", async () => {
    vi.mocked(archiveApi.createContractCard).mockRejectedValue(
      new ApiError("Номер 0786/1/25 уже есть", 400, "Номер 0786/1/25 уже есть"),
    );
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "Секлюц" } });
    fireEvent.change(screen.getByLabelText("Номер договора"), { target: { value: "0786/1/25" } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));

    expect(await screen.findByRole("alert")).toHaveTextContent("Номер 0786/1/25 уже есть");
    expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Номер договора")).toHaveValue("0786/1/25");
  });

  it("creates a card from the first dropped docx and opens it with the parsed blank", async () => {
    const docx = new File(["card"], "Секлюц.docx");
    const extra = new File(["other"], "second.xlsx");
    vi.mocked(archiveApi.parseContractCardFile).mockResolvedValue({
      fields: {
        short_name: "Секлюц",
        full_name: "ИП Секлюцкий Иван",
        legal_form: "ip",
        inn: "440117229007",
        bank_name: "Альфа",
        account: "40702810000000000007",
      },
      doubtful: ["inn"],
      accounts: ["40702810000000000007", "40702810900000000009"],
      banks: [
        {
          bank_name: "Альфа",
          account: "40702810000000000007",
          corr_account: "30101810000000000760",
          bik: "044525225",
        },
        {
          bank_name: "ВТБ",
          account: "40702810900000000009",
          corr_account: "30101810400000000225",
          bik: "044525974",
        },
      ],
      source_text: "Текст карточки",
      verify_failed: false,
    });
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(
      card({
        id: 11,
        name: "Секлюц",
        imported_name: "Секлюц",
        inn: "440117229007",
        number: "1034/10/26",
        contract_date: "2026-10-21",
      }),
    );
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    expect(within(form).getByText("или перетащите файл сюда")).toBeInTheDocument();
    expect(within(form).getByLabelText("Файл новой карточки")).toHaveAttribute(
      "accept",
      ".docx,.pdf,.xlsx,image/jpeg,image/png,application/pdf",
    );
    fireEvent.change(within(form).getByLabelText("Номер договора"), {
      target: { value: "0001/10/26" },
    });
    fireEvent.drop(within(form).getByLabelText("Зона файла новой карточки"), {
      dataTransfer: { files: [docx, extra] },
    });

    await waitFor(() => {
      expect(archiveApi.createContractCard).toHaveBeenCalledTimes(1);
    });
    expect(archiveApi.parseContractCardFile).toHaveBeenCalledTimes(1);
    expect(archiveApi.parseContractCardFile).toHaveBeenCalledWith(docx);
    expect(archiveApi.createContractCard).toHaveBeenCalledWith({
      name: "Секлюц",
      inn: "440117229007",
    });
    expect(archiveApi.findDuplicateContractCards).toHaveBeenCalledWith("Секлюц", "440117229007");
    expect(archiveApi.parseContractCardFile.mock.invocationCallOrder[0]).toBeLessThan(
      archiveApi.findDuplicateContractCards.mock.invocationCallOrder[0],
    );
    expect(archiveApi.findDuplicateContractCards.mock.invocationCallOrder[0]).toBeLessThan(
      archiveApi.createContractCard.mock.invocationCallOrder[0],
    );
    expect(screen.queryByRole("heading", { name: "Такая карточка уже есть" })).not.toBeInTheDocument();
    expect(archiveApi.createContractCard.mock.calls[0]?.[0]).not.toHaveProperty("number");
    expect(archiveApi.parseContractCard).not.toHaveBeenCalled();
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();

    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("Секлюц");
    expect(opened).toHaveTextContent("1034/10/26");
    expect(opened).not.toHaveTextContent("0001/10/26");
    expect(within(opened).getByLabelText("ИНН контрагента")).toHaveValue("440117229007");
    const blank = within(opened).getByRole("form", { name: "Реквизиты договора" });
    expect(within(blank).getByLabelText("Краткое наименование")).toHaveValue("Секлюц");
    expect(within(blank).getByLabelText("ИНН").closest("[data-doubtful='true']")).toBeTruthy();
    expect(within(blank).getByLabelText("Банк")).toHaveValue("");
    expect(within(blank).getByRole("button", { name: "ВТБ 40702810900000000009" })).toBeInTheDocument();
    expect(within(blank).getByLabelText("Текст карточки")).toHaveTextContent("Текст карточки");
    expect(screen.getByText("Проверьте по карточке: ИНН.")).toBeInTheDocument();
    expect(screen.getByLabelText("Номер договора")).toHaveValue("");
  });

  it("uses the full name when the dropped card has no short name", async () => {
    vi.mocked(archiveApi.parseContractCardFile).mockResolvedValue({
      fields: { short_name: "   ", full_name: "ООО Ромашка", inn: "7701000001" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(
      card({ id: 12, name: "ООО Ромашка", imported_name: "ООО Ромашка", inn: "7701000001", number: "1035/10/26" }),
    );
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.change(within(form).getByLabelText("Файл новой карточки"), {
      target: { files: [new File(["card"], "ignored.docx")] },
    });

    await waitFor(() => {
      expect(archiveApi.createContractCard).toHaveBeenCalledWith({
        name: "ООО Ромашка",
        inn: "7701000001",
      });
    });
    expect(await screen.findByRole("heading", { name: "ООО Ромашка" })).toBeInTheDocument();
    expect(archiveApi.createContractCard.mock.calls[0]?.[0]).not.toHaveProperty("number");
  });

  it("names the card from the file when the parse has no short or full name", async () => {
    vi.mocked(archiveApi.parseContractCardFile).mockResolvedValue({
      fields: { inn: "7707083893" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(
      card({ id: 13, name: "Ромашка", imported_name: "Ромашка", inn: "7707083893", number: "1036/10/26" }),
    );
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.drop(within(form).getByLabelText("Зона файла новой карточки"), {
      dataTransfer: { files: [new File(["card"], "Ромашка.docx")] },
    });

    await waitFor(() => {
      expect(archiveApi.createContractCard).toHaveBeenCalledWith({
        name: "Ромашка",
        inn: "7707083893",
      });
    });
    expect(await screen.findByRole("heading", { name: "Ромашка" })).toBeInTheDocument();
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
  });

  it("shows a parse refusal on the new-card form and does not create a row", async () => {
    vi.mocked(archiveApi.parseContractCardFile).mockRejectedValue(new Error("это не карточка контрагента"));
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.change(within(form).getByLabelText("Имя"), { target: { value: "Черновик" } });
    fireEvent.drop(within(form).getByLabelText("Зона файла новой карточки"), {
      dataTransfer: { files: [new File(["passport"], "passport.docx"), new File(["x"], "second.pdf")] },
    });

    expect(await within(form).findByRole("alert")).toHaveTextContent("это не карточка контрагента");
    expect(archiveApi.parseContractCardFile).toHaveBeenCalledTimes(1);
    expect(archiveApi.findDuplicateContractCards).not.toHaveBeenCalled();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(archiveApi.parseContractCard).not.toHaveBeenCalled();
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
    expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
    expect(within(form).getByLabelText("Имя")).toHaveValue("Черновик");
  });

  it("closes the opened card after УДАЛИТЬ and drops it from the saved list", async () => {
    const saved = card({ id: 1, number: "0793/02/25" });
    let searches = 0;
    vi.mocked(archiveApi.searchContractCards).mockImplementation(async () => {
      searches += 1;
      if (searches === 1) {
        return { items: [saved], count: 1 };
      }
      return { items: [], count: 0 };
    });
    vi.mocked(archiveApi.deleteContractCard).mockResolvedValue(undefined);
    renderList();

    fireEvent.click(await screen.findByRole("button", { name: /0793\/02\/25/ }));
    expect(screen.getByRole("region", { name: "Карточка договора" })).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Удалить карточку" }));
    const confirm = screen.getByRole("button", { name: "Подтвердить удаление" });
    fireEvent.click(confirm);
    fireEvent.change(screen.getByLabelText("Подтверждение"), { target: { value: "удалить" } });
    fireEvent.click(confirm);
    expect(archiveApi.deleteContractCard).not.toHaveBeenCalled();

    fireEvent.change(screen.getByLabelText("Подтверждение"), { target: { value: "УДАЛИТЬ" } });
    fireEvent.click(confirm);

    await waitFor(() => {
      expect(archiveApi.deleteContractCard).toHaveBeenCalledWith(1);
    });
    await waitFor(() => {
      expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
      expect(screen.queryByText("0793/02/25")).not.toBeInTheDocument();
    });
    expect(searches).toBeGreaterThanOrEqual(2);
  });

  it("shows a duplicate dialog and opens the found card without creating", async () => {
    const found = card({
      id: 21,
      name: "ООО Ромашка",
      imported_name: "ООО Ромашка",
      inn: "7701000001",
      number: "0001/10/26",
    });
    vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({ items: [found], count: 1 });
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "ооо  ромашка" } });
    fireEvent.change(screen.getByLabelText("ИНН"), { target: { value: " 7702000002 " } });
    fireEvent.change(screen.getByLabelText("Номер договора"), { target: { value: "0009/10/26" } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));

    const dialog = await screen.findByRole("dialog");
    expect(within(dialog).getByRole("heading", { name: "Такая карточка уже есть" })).toBeInTheDocument();
    expect(dialog).toHaveTextContent("ООО Ромашка");
    expect(dialog).toHaveTextContent("7701000001");
    expect(dialog).toHaveTextContent("0001/10/26");
    expect(archiveApi.findDuplicateContractCards).toHaveBeenCalledWith("ооо  ромашка", "7702000002");
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();

    fireEvent.click(within(dialog).getByRole("button", { name: "Редактировать" }));

    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("0001/10/26");
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Имя")).toHaveValue("ооо  ромашка");
    expect(screen.getByLabelText("Номер договора")).toHaveValue("0009/10/26");
  });

  it("creates the typed card once when the manager confirms the duplicate", async () => {
    vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({
      items: [
        card({
          id: 21,
          name: "ООО Ромашка",
          imported_name: "ООО Ромашка",
          inn: "7701000001",
          number: "0001/10/26",
        }),
      ],
      count: 1,
    });
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(
      card({
        id: 30,
        name: "ооо  ромашка",
        imported_name: "ооо  ромашка",
        inn: "7702000002",
        number: "0009/10/26",
      }),
    );
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "ооо  ромашка" } });
    fireEvent.change(screen.getByLabelText("ИНН"), { target: { value: "7702000002" } });
    fireEvent.change(screen.getByLabelText("Номер договора"), { target: { value: "0009/10/26" } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));
    fireEvent.click(await screen.findByRole("button", { name: "Всё равно создать" }));

    await waitFor(() => {
      expect(archiveApi.createContractCard).toHaveBeenCalledTimes(1);
    });
    expect(archiveApi.createContractCard).toHaveBeenCalledWith({
      name: "ооо  ромашка",
      inn: "7702000002",
      number: "0009/10/26",
    });
    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("0009/10/26");
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    const form = screen.getByRole("form", { name: "Новая карточка" });
    expect(within(form).getByLabelText("Имя")).toHaveValue("");
    expect(within(form).getByLabelText("ИНН")).toHaveValue("");
    expect(within(form).getByLabelText("Номер договора")).toHaveValue("");
  });

  it("opens the second duplicate and does not create a card", async () => {
    vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({
      items: [
        card({ id: 21, name: "ООО Ромашка", number: "0001/10/26", inn: "7701000001" }),
        card({ id: 22, name: "Другая", number: "0002/10/26", inn: null }),
      ],
      count: 2,
    });
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "Другая" } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));

    const dialog = await screen.findByRole("dialog");
    const rows = within(dialog).getAllByRole("listitem");
    expect(rows).toHaveLength(2);
    expect(within(dialog).getAllByRole("button", { name: "Всё равно создать" })).toHaveLength(1);
    expect(dialog).toHaveTextContent("—");
    expect(dialog).toHaveTextContent("0002/10/26");
    fireEvent.click(within(rows[1]).getByRole("button", { name: "Редактировать" }));

    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("0002/10/26");
    expect(opened).not.toHaveTextContent("0001/10/26");
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });

  it("does not create a dropped card until the manager confirms, then opens the parsed blank", async () => {
    const docx = new File(["card"], "Секлюц.docx");
    vi.mocked(archiveApi.parseContractCardFile).mockResolvedValue({
      fields: { short_name: "Секлюц", inn: "440117229007" },
      doubtful: ["inn"],
      accounts: [],
      verify_failed: false,
    });
    vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({
      items: [card({ id: 21, name: "Секлюц", number: "0001/10/26", inn: "440117229007" })],
      count: 1,
    });
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(
      card({
        id: 31,
        name: "Секлюц",
        imported_name: "Секлюц",
        inn: "440117229007",
        number: "1034/10/26",
        short_name: null,
      }),
    );
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.change(within(form).getByLabelText("Номер договора"), { target: { value: "0001/10/26" } });
    fireEvent.drop(within(form).getByLabelText("Зона файла новой карточки"), {
      dataTransfer: { files: [docx] },
    });

    const dialog = await screen.findByRole("dialog");
    expect(dialog).toHaveTextContent("0001/10/26");
    expect(archiveApi.parseContractCardFile).toHaveBeenCalledWith(docx);
    expect(archiveApi.findDuplicateContractCards).toHaveBeenCalledWith("Секлюц", "440117229007");
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(archiveApi.parseContractCardFile.mock.invocationCallOrder[0]).toBeLessThan(
      archiveApi.findDuplicateContractCards.mock.invocationCallOrder[0],
    );

    fireEvent.click(within(dialog).getByRole("button", { name: "Всё равно создать" }));

    await waitFor(() => {
      expect(archiveApi.createContractCard).toHaveBeenCalledTimes(1);
    });
    expect(archiveApi.createContractCard).toHaveBeenCalledWith({
      name: "Секлюц",
      inn: "440117229007",
    });
    expect(archiveApi.createContractCard.mock.calls[0]?.[0]).not.toHaveProperty("number");
    expect(archiveApi.findDuplicateContractCards.mock.invocationCallOrder[0]).toBeLessThan(
      archiveApi.createContractCard.mock.invocationCallOrder[0],
    );
    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    const blank = within(opened).getByRole("form", { name: "Реквизиты договора" });
    expect(within(blank).getByLabelText("Краткое наименование")).toHaveValue("Секлюц");
    expect(within(blank).getByLabelText("ИНН")).toHaveValue("440117229007");
    expect(within(blank).getByLabelText("ИНН").closest("[data-doubtful='true']")).toBeTruthy();
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Номер договора")).toHaveValue("");
  });

  it("opens the found card from a dropped file without the parsed blank", async () => {
    vi.mocked(archiveApi.parseContractCardFile).mockResolvedValue({
      fields: { short_name: "Из файла", full_name: "Не подставлять", inn: "440117229007" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({
      items: [
        card({
          id: 21,
          name: "Из файла",
          imported_name: "Из файла",
          inn: "7701000001",
          short_name: null,
          number: "0001/10/26",
        }),
      ],
      count: 1,
    });
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.drop(within(form).getByLabelText("Зона файла новой карточки"), {
      dataTransfer: { files: [new File(["card"], "card.docx")] },
    });

    const dialog = await screen.findByRole("dialog");
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    fireEvent.click(within(dialog).getByRole("button", { name: "Редактировать" }));

    const opened = await screen.findByRole("region", { name: "Карточка договора" });
    expect(opened).toHaveTextContent("0001/10/26");
    const blank = within(opened).getByRole("form", { name: "Реквизиты договора" });
    expect(within(blank).getByLabelText("Краткое наименование")).toHaveValue("");
    expect(within(blank).getByLabelText("ИНН")).toHaveValue("7701000001");
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(archiveApi.parseContractCard).not.toHaveBeenCalled();
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
  });

  it("asks again after the dialog is closed without creating a card", async () => {
    vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({
      items: [card({ id: 21, name: "Ромашка", number: "0001/10/26", inn: "7701000001" })],
      count: 1,
    });
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "Ромашка" } });
    fireEvent.change(screen.getByLabelText("ИНН"), { target: { value: "7701000001" } });
    fireEvent.change(screen.getByLabelText("Номер договора"), { target: { value: "0008/10/26" } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));

    const dialog = await screen.findByRole("dialog");
    fireEvent.click(within(dialog).getByRole("button", { name: "Закрыть" }));

    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(screen.getByLabelText("Имя")).toHaveValue("Ромашка");
    expect(screen.getByLabelText("ИНН")).toHaveValue("7701000001");
    expect(screen.getByLabelText("Номер договора")).toHaveValue("0008/10/26");

    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));
    expect(await screen.findByRole("heading", { name: "Такая карточка уже есть" })).toBeInTheDocument();
    expect(archiveApi.findDuplicateContractCards).toHaveBeenCalledTimes(2);
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });

  it("does not search duplicates when the parsed card has no name", async () => {
    vi.mocked(archiveApi.parseContractCardFile).mockResolvedValue({
      fields: { short_name: "  ", full_name: "" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.drop(within(form).getByLabelText("Зона файла новой карточки"), {
      dataTransfer: { files: [new File(["card"], " .docx")] },
    });

    expect(await within(form).findByRole("alert")).toHaveTextContent("Укажите имя контрагента");
    expect(archiveApi.findDuplicateContractCards).not.toHaveBeenCalled();
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it("keeps the typed number when create anyway hits a taken number", async () => {
    vi.mocked(archiveApi.findDuplicateContractCards).mockResolvedValue({
      items: [card({ id: 21, name: "Секлюц", number: "0001/10/26", inn: "440117229007" })],
      count: 1,
    });
    vi.mocked(archiveApi.createContractCard).mockRejectedValue(
      new ApiError("Номер 0786/1/25 уже есть", 400, "Номер 0786/1/25 уже есть"),
    );
    renderList();

    fireEvent.change(screen.getByLabelText("Имя"), { target: { value: "Секлюц" } });
    fireEvent.change(screen.getByLabelText("Номер договора"), { target: { value: "0786/1/25" } });
    fireEvent.click(screen.getByRole("button", { name: "Создать карточку" }));
    fireEvent.click(await screen.findByRole("button", { name: "Всё равно создать" }));

    expect(await screen.findByRole("alert")).toHaveTextContent("Номер 0786/1/25 уже есть");
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Номер договора")).toHaveValue("0786/1/25");
    expect(screen.getByLabelText("Имя")).toHaveValue("Секлюц");
    expect(archiveApi.createContractCard).toHaveBeenCalledTimes(1);
    expect(archiveApi.createContractCard).toHaveBeenCalledWith({
      name: "Секлюц",
      number: "0786/1/25",
    });
  });

  it("shows a duplicate search server error and does not create a card", async () => {
    vi.mocked(archiveApi.findDuplicateContractCards).mockRejectedValue(
      new ApiError(
        "Внутренняя ошибка сервера. Повторите попытку позже.",
        500,
        "Внутренняя ошибка сервера. Повторите попытку позже.",
      ),
    );
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.change(within(form).getByLabelText("Имя"), { target: { value: "ООО Ромашка" } });
    fireEvent.change(within(form).getByLabelText("ИНН"), { target: { value: "7701000001" } });
    fireEvent.change(within(form).getByLabelText("Номер договора"), { target: { value: "0009/10/26" } });
    fireEvent.click(within(form).getByRole("button", { name: "Создать карточку" }));

    expect(await within(form).findByRole("alert")).toHaveTextContent(
      "Внутренняя ошибка сервера. Повторите попытку позже.",
    );
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
    expect(within(form).getByLabelText("Имя")).toHaveValue("ООО Ромашка");
    expect(within(form).getByLabelText("ИНН")).toHaveValue("7701000001");
    expect(within(form).getByLabelText("Номер договора")).toHaveValue("0009/10/26");
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });

  it("shows a duplicate search network error and does not create a card", async () => {
    vi.mocked(archiveApi.findDuplicateContractCards).mockRejectedValue(
      new ApiError(
        "Сервер недоступен. Подождите пару секунд после запуска и обновите страницу.",
        0,
        "Сервер недоступен. Подождите пару секунд после запуска и обновите страницу.",
      ),
    );
    renderList();

    const form = screen.getByRole("form", { name: "Новая карточка" });
    fireEvent.change(within(form).getByLabelText("Имя"), { target: { value: "ООО Ромашка" } });
    fireEvent.click(within(form).getByRole("button", { name: "Создать карточку" }));

    expect(await within(form).findByRole("alert")).toHaveTextContent(
      "Сервер недоступен. Подождите пару секунд после запуска и обновите страницу.",
    );
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    expect(screen.queryByRole("region", { name: "Карточка договора" })).not.toBeInTheDocument();
    expect(within(form).getByLabelText("Имя")).toHaveValue("ООО Ромашка");
    expect(archiveApi.createContractCard).not.toHaveBeenCalled();
  });
});
