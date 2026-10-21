import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { ContractCardDrawer } from "@/features/commercial-archive/components/ContractCardDrawer";
import type { ContractCard } from "@/features/commercial-archive/types/supplyContract";
import { ApiError } from "@/shared/lib/apiError";

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    createSupplyContract: vi.fn(),
    getSupplyContract: vi.fn(),
    listSupplyContracts: vi.fn(),
    parseSupplyContract: vi.fn(),
    downloadSupplyContract: vi.fn(),
    downloadEdoAgreement: vi.fn(),
    replaceSupplyContract: vi.fn(),
    patchContractCard: vi.fn(),
    patchSupplyContract: vi.fn(),
    saveContractRequisites: vi.fn(),
    parseContractCard: vi.fn(),
    downloadContractDocument: vi.fn(),
    downloadContractEdoAgreement: vi.fn(),
    getContractCard: vi.fn(),
    deleteContractCard: vi.fn(),
  },
}));

vi.mock("@/shared/lib/downloadFile", () => ({
  saveBlobAs: vi.fn(),
}));

const card = (overrides: Partial<ContractCard> = {}): ContractCard => ({
  id: 9,
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

const downloadResult = {
  blob: new Blob(["file"]),
  filename: "dogovor-9.docx",
  contentType: "application/octet-stream",
};

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

const renderDrawer = (value: ContractCard = card()) => wrap(<ContractCardDrawer card={value} />);

const fillBlank = () => {
  fireEvent.change(screen.getByLabelText("Вид"), { target: { value: "ooo" } });
  fireEvent.change(screen.getByLabelText("Полное наименование"), { target: { value: "ООО Секлюц" } });
  fireEvent.change(screen.getByLabelText("Краткое наименование"), { target: { value: "Секлюц" } });
  fireEvent.change(screen.getByLabelText("ФИО подписанта"), { target: { value: "Иванов Иван" } });
  fireEvent.change(screen.getByLabelText("ОГРН"), { target: { value: "1024400000000" } });
  fireEvent.change(screen.getByLabelText("Юридический адрес"), { target: { value: "Кострома" } });
  fireEvent.change(screen.getByLabelText("E-mail"), { target: { value: "a@example.com" } });
  fireEvent.change(screen.getByLabelText("Банк"), { target: { value: "Банк" } });
  fireEvent.change(screen.getByLabelText("Расчётный счёт"), { target: { value: "40702810000000000001" } });
  fireEvent.change(screen.getByLabelText("Корсчёт"), { target: { value: "30101810000000000001" } });
  fireEvent.change(screen.getByLabelText("БИК"), { target: { value: "044525225" } });
};

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("ContractCardDrawer", () => {
  it("shows the name, INN and number and does not create a contract from a KP", () => {
    renderDrawer();

    expect(screen.getByLabelText("Имя контрагента")).toHaveValue("Секлюцкий");
    expect(screen.getByLabelText("ИНН контрагента")).toHaveValue("440117229007");
    expect(screen.getByLabelText("Номер")).toHaveValue("0793/02/25");
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    expect(screen.getByLabelText("Вид")).toHaveValue("");
    expect(screen.getByLabelText("Основание")).toHaveValue("устав");
    expect(screen.queryByRole("region", { name: "Реестр договоров" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "Договор поставки" })).not.toBeInTheDocument();
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();
    expect(archiveApi.getSupplyContract).not.toHaveBeenCalled();
    expect(archiveApi.listSupplyContracts).not.toHaveBeenCalled();
    expect(archiveApi.getContractCard).not.toHaveBeenCalled();
  });

  it("keeps requisites save disabled until a legal form is chosen", () => {
    renderDrawer();

    const save = screen.getByRole("button", { name: "Сохранить реквизиты" });
    const kind = screen.getByLabelText("Вид") as HTMLSelectElement;
    expect(kind.value).toBe("");
    expect(kind.selectedOptions[0]?.textContent).toBe("Выберите вид");
    expect(save).toBeDisabled();
    expect(save).toHaveAttribute("title", "Выберите вид");

    fireEvent.submit(screen.getByRole("form", { name: "Реквизиты договора" }));
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();

    fireEvent.change(kind, { target: { value: "ooo" } });
    expect(save).toBeEnabled();
    expect(save).not.toHaveAttribute("title");
  });

  it("shows requisites already stored on the card and enables save", () => {
    renderDrawer(
      card({
        legal_form: "ooo",
        full_name: "ООО Ромашка",
        short_name: "Ромашка",
        signatory_verb: "действующей",
        authority_basis: "устав",
      }),
    );

    expect(screen.getByLabelText("Вид")).toHaveValue("ooo");
    expect(screen.queryByRole("option", { name: "Выберите вид" })).not.toBeInTheDocument();
    expect(screen.getByLabelText("Полное наименование")).toHaveValue("ООО Ромашка");
    expect(screen.getByLabelText("Краткое наименование")).toHaveValue("Ромашка");
    expect(screen.getByLabelText("Род")).toHaveValue("действующей");
    expect(screen.getByRole("button", { name: "Сохранить реквизиты" })).toBeEnabled();
  });

  it("asks for the power of attorney number and date before saving requisites", () => {
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Вид"), { target: { value: "ip" } });
    fireEvent.change(screen.getByLabelText("Основание"), { target: { value: "доверенность" } });

    const save = screen.getByRole("button", { name: "Сохранить реквизиты" });
    expect(save).toBeDisabled();
    expect(save).toHaveAttribute("title", "Укажите номер и дату доверенности");
    fireEvent.submit(screen.getByRole("form", { name: "Реквизиты договора" }));
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();

    fireEvent.change(screen.getByLabelText("Номер доверенности"), { target: { value: "15" } });
    fireEvent.change(screen.getByLabelText("Дата доверенности"), { target: { value: "01.02.2024" } });
    expect(save).toBeEnabled();
  });

  it("shows не полные until the saved requisites come back complete", async () => {
    vi.mocked(archiveApi.saveContractRequisites).mockResolvedValue({
      id: 9,
      counterparty_id: null,
      number: "0793/02/25",
      contract_date: "2025-02-01",
      manager_name: "ivan",
      status: "нет",
      legal_form: "ooo",
      full_name: "ООО Секлюц",
      short_name: "Секлюц",
      signatory_name: "Иванов Иван",
      email: "a@example.com",
      requisites_complete: true,
    });
    renderDrawer(card({ requisites_complete: false }));

    expect(screen.getByText("не полные")).toBeInTheDocument();

    fillBlank();
    fireEvent.click(screen.getByRole("button", { name: "Сохранить реквизиты" }));

    await waitFor(() => {
      expect(screen.queryByText("не полные")).not.toBeInTheDocument();
    });
    expect(archiveApi.saveContractRequisites).toHaveBeenCalledTimes(1);
  });

  it("hides не полные when the card requisites are already complete", () => {
    renderDrawer(card({ requisites_complete: true }));

    expect(screen.queryByText("не полные")).not.toBeInTheDocument();
  });

  it("saves the blank through requisites and does not create a supply contract", async () => {
    vi.mocked(archiveApi.saveContractRequisites).mockResolvedValue({
      id: 9,
      counterparty_id: null,
      number: "0793/02/25",
      contract_date: "2025-02-01",
      manager_name: "ivan",
      status: "нет",
      legal_form: "ooo",
      full_name: "ООО Секлюц",
      short_name: "Секлюц",
      signatory_name: "Иванов Иван",
      email: "a@example.com",
    });
    renderDrawer();

    fillBlank();
    fireEvent.click(screen.getByRole("button", { name: "Сохранить реквизиты" }));

    await waitFor(() => {
      expect(archiveApi.saveContractRequisites).toHaveBeenCalledWith(
        9,
        expect.objectContaining({
          legal_form: "ooo",
          full_name: "ООО Секлюц",
          short_name: "Секлюц",
          signatory_name: "Иванов Иван",
          inn: "440117229007",
          ogrn: "1024400000000",
          legal_address: "Кострома",
          email: "a@example.com",
          bank_name: "Банк",
          account: "40702810000000000001",
          corr_account: "30101810000000000001",
          bik: "044525225",
        }),
      );
    });
    expect(archiveApi.saveContractRequisites.mock.calls[0]?.[1]).not.toHaveProperty("contract_date");
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    expect(screen.queryByText("01.10.2026")).not.toBeInTheDocument();
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();
    expect(archiveApi.replaceSupplyContract).not.toHaveBeenCalled();
  });

  it("shows the server refusal when Word is requested before requisites are filled", async () => {
    vi.mocked(archiveApi.downloadContractDocument).mockRejectedValue(
      new ApiError(
        "Сначала заполните реквизиты договора",
        400,
        "Сначала заполните реквизиты договора",
      ),
    );
    renderDrawer();

    fireEvent.click(screen.getByRole("button", { name: "Договор Word" }));

    expect(await screen.findByRole("alert")).toHaveTextContent("Сначала заполните реквизиты договора");
    expect(archiveApi.downloadContractDocument).toHaveBeenCalledWith(9, "docx");
    expect(archiveApi.downloadSupplyContract).not.toHaveBeenCalled();
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();
  });

  it("downloads PDF and the EDO agreement by contract id", async () => {
    vi.mocked(archiveApi.downloadContractDocument).mockResolvedValue({
      ...downloadResult,
      filename: "dogovor-9.pdf",
    });
    vi.mocked(archiveApi.downloadContractEdoAgreement).mockResolvedValue({
      ...downloadResult,
      filename: "soglashenie-edo-9.docx",
    });
    renderDrawer();

    fireEvent.click(screen.getByRole("button", { name: "Договор PDF" }));
    fireEvent.click(screen.getByRole("button", { name: "Соглашение об ЭДО Word" }));
    fireEvent.click(screen.getByRole("button", { name: "Соглашение об ЭДО PDF" }));

    await waitFor(() => {
      expect(archiveApi.downloadContractDocument).toHaveBeenCalledWith(9, "pdf");
      expect(archiveApi.downloadContractEdoAgreement).toHaveBeenCalledWith(9, "docx");
      expect(archiveApi.downloadContractEdoAgreement).toHaveBeenCalledWith(9, "pdf");
    });
    expect(archiveApi.downloadSupplyContract).not.toHaveBeenCalled();
    expect(archiveApi.downloadEdoAgreement).not.toHaveBeenCalled();
  });

  it("patches the status by contract id", async () => {
    vi.mocked(archiveApi.patchSupplyContract).mockResolvedValue({
      id: 9,
      counterparty_id: null,
      number: "0793/02/25",
      contract_date: "2025-02-01",
      manager_name: null,
      status: "подписан по ЭДО",
      legal_form: null,
      full_name: null,
      short_name: null,
      signatory_name: null,
      email: null,
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Статус"), { target: { value: "подписан по ЭДО" } });

    await waitFor(() => {
      expect(archiveApi.patchSupplyContract).toHaveBeenCalledWith(9, { status: "подписан по ЭДО" });
      expect(screen.getByLabelText("Статус")).toHaveValue("подписан по ЭДО");
    });
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();
  });

  it("keeps the previous status when the patch fails", async () => {
    vi.mocked(archiveApi.patchSupplyContract).mockRejectedValue(
      new ApiError("Статус не сохранился", 400, "Статус не сохранился"),
    );
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Статус"), { target: { value: "оригинал в бухгалтерии" } });

    expect(await screen.findByRole("alert")).toHaveTextContent("Статус не сохранился");
    expect(screen.getByLabelText("Статус")).toHaveValue("нет");
  });

  it("patches the scan note by contract id", async () => {
    vi.mocked(archiveApi.patchSupplyContract).mockResolvedValue({
      id: 9,
      counterparty_id: null,
      number: "0793/02/25",
      contract_date: "2025-02-01",
      manager_name: null,
      status: "нет",
      scan_note: "получен скан",
      has_scan: false,
      legal_form: null,
      full_name: null,
      short_name: null,
      signatory_name: null,
      email: null,
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Отметка скана"), { target: { value: "получен скан" } });

    await waitFor(() => {
      expect(archiveApi.patchSupplyContract).toHaveBeenCalledWith(9, { scan_note: "получен скан" });
      expect(screen.getByLabelText("Отметка скана")).toHaveValue("получен скан");
    });
    expect(screen.getByText("файла нет")).toBeInTheDocument();
  });

  it("patches the scan file by contract id and shows that the file is stored", async () => {
    const file = new File(["scan"], "scan.pdf", { type: "application/pdf" });
    vi.mocked(archiveApi.patchSupplyContract).mockResolvedValue({
      id: 9,
      counterparty_id: null,
      number: "0793/02/25",
      contract_date: "2025-02-01",
      manager_name: null,
      status: "нет",
      scan_note: null,
      has_scan: true,
      legal_form: null,
      full_name: null,
      short_name: null,
      signatory_name: null,
      email: null,
    });
    renderDrawer();

    expect(screen.getByText("файла нет")).toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Файл скана"), { target: { files: [file] } });

    await waitFor(() => {
      expect(archiveApi.patchSupplyContract).toHaveBeenCalledWith(9, { file });
    });
    expect(await screen.findByText("файл есть")).toBeInTheDocument();
  });

  it("parses a card file by contract id and fills the blank", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { full_name: "ООО Секлюц", legal_form: "ooo", inn: "440117229007" },
      doubtful: ["inn"],
      accounts: [],
      verify_failed: false,
    });
    renderDrawer();

    const file = new File(["card"], "card.png", { type: "image/png" });
    fireEvent.change(screen.getByLabelText("Файл карточки"), { target: { files: [file] } });

    await waitFor(() => {
      expect(archiveApi.parseContractCard).toHaveBeenCalledWith(9, file);
    });
    expect(screen.getByLabelText("Полное наименование")).toHaveValue("ООО Секлюц");
    expect(screen.getByLabelText("Вид")).toHaveValue("ooo");
    expect(screen.getByLabelText("ИНН").closest("[data-doubtful='true']")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Сохранить реквизиты" })).toBeEnabled();
    expect(archiveApi.parseSupplyContract).not.toHaveBeenCalled();
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();
  });

  it("keeps a valid basis when the file returns wording outside the list", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { authority_basis: "Устава" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    await waitFor(() => {
      expect(archiveApi.parseContractCard).toHaveBeenCalled();
    });
    expect(screen.getByLabelText("Основание")).toHaveValue("устав");
  });

  it("lets the manager choose one bank when the file lists two", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: {
        legal_address: "ул. Ленина, 1",
        signatory_name: "Иванов Иван Иванович",
        bank_name: "Альфа",
        account: "40702810000000000007",
      },
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
      doubtful: ["legal_address", "signatory_name", "bank_name", "account", "corr_account", "bik"],
      verify_failed: false,
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    expect(await screen.findByText("Проверьте по карточке: юридический адрес, ФИО.")).toBeInTheDocument();
    expect(screen.getByLabelText("Банк")).toHaveValue("");
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("");
    expect(screen.getByLabelText("Корсчёт")).toHaveValue("");
    expect(screen.getByLabelText("БИК")).toHaveValue("");

    fireEvent.click(screen.getByRole("button", { name: "ВТБ 40702810900000000009" }));

    expect(screen.getByLabelText("Банк")).toHaveValue("ВТБ");
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("40702810900000000009");
    expect(screen.getByLabelText("Корсчёт")).toHaveValue("30101810400000000225");
    expect(screen.getByLabelText("БИК")).toHaveValue("044525974");
    expect(screen.getByLabelText("Банк").closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("Юридический адрес").closest("[data-doubtful='true']")).toBeTruthy();
  });

  it("writes only the chosen settlement account when the file lists two", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { account: "40702810000000000007" },
      accounts: ["40702810000000000007", "40802810000000000015"],
      doubtful: ["account"],
      verify_failed: false,
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    expect(await screen.findByRole("button", { name: "40802810000000000015" })).toBeInTheDocument();
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("");
    fireEvent.click(screen.getByRole("button", { name: "40802810000000000015" }));
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("40802810000000000015");
  });

  it("patches the name, INN and number without rewriting the date on screen", async () => {
    vi.mocked(archiveApi.patchContractCard).mockResolvedValue(
      card({
        name: "Новое имя",
        inn: "7707083893",
        number: "1035/10/26",
        contract_date: "2025-02-01",
      }),
    );
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Номер"), { target: { value: "1035/10/26" } });
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    expect(screen.queryByText("01.10.2026")).not.toBeInTheDocument();

    fireEvent.change(screen.getByLabelText("Имя контрагента"), { target: { value: "  Новое имя  " } });
    fireEvent.change(screen.getByLabelText("ИНН контрагента"), { target: { value: " 7707083893 " } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить карточку" }));

    await waitFor(() => {
      expect(archiveApi.patchContractCard).toHaveBeenCalledWith(9, {
        name: "Новое имя",
        inn: "7707083893",
        number: "1035/10/26",
      });
    });
    expect(archiveApi.patchContractCard.mock.calls[0]?.[1]).not.toHaveProperty("contract_date");
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    expect(screen.getByLabelText("Номер")).toHaveValue("1035/10/26");
    expect(archiveApi.createSupplyContract).not.toHaveBeenCalled();
  });

  it("shows the contract date from the server response when that date changed", async () => {
    vi.mocked(archiveApi.patchContractCard).mockResolvedValue(
      card({ number: "1035/10/26", contract_date: "2026-03-15" }),
    );
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Номер"), { target: { value: "1035/10/26" } });
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Сохранить карточку" }));

    expect(await screen.findByText("Дата: 15.03.2026")).toBeInTheDocument();
    expect(screen.queryByText("Дата: 01.02.2025")).not.toBeInTheDocument();
    expect(screen.queryByText("01.10.2026")).not.toBeInTheDocument();
  });

  it("does not patch a card with a blank name or number", () => {
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Имя контрагента"), { target: { value: "   " } });
    expect(screen.getByRole("button", { name: "Сохранить карточку" })).toBeDisabled();
    fireEvent.submit(screen.getByRole("form", { name: "Данные карточки" }));
    expect(archiveApi.patchContractCard).not.toHaveBeenCalled();

    fireEvent.change(screen.getByLabelText("Имя контрагента"), { target: { value: "Секлюцкий" } });
    fireEvent.change(screen.getByLabelText("Номер"), { target: { value: "  " } });
    expect(screen.getByRole("button", { name: "Сохранить карточку" })).toBeDisabled();
    fireEvent.submit(screen.getByRole("form", { name: "Данные карточки" }));
    expect(archiveApi.patchContractCard).not.toHaveBeenCalled();
  });

  it("shows a duplicate number error and leaves the date as it was", async () => {
    vi.mocked(archiveApi.patchContractCard).mockRejectedValue(
      new ApiError("Номер 1035/10/26 уже есть", 400, "Номер 1035/10/26 уже есть"),
    );
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Номер"), { target: { value: "1035/10/26" } });
    fireEvent.click(screen.getByRole("button", { name: "Сохранить карточку" }));

    expect(await screen.findByRole("alert")).toHaveTextContent("Номер 1035/10/26 уже есть");
    expect(screen.getByText("Дата: 01.02.2025")).toBeInTheDocument();
    expect(screen.getByLabelText("Номер")).toHaveValue("1035/10/26");
  });

  it("parses only the first dropped file for docx, xlsx and pdf", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { inn: "440117229007" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
      source_text: "Текст карточки",
    });
    renderDrawer();

    const zone = screen.getByLabelText("Зона файла карточки");
    const docx = new File(["one"], "first.docx");
    const xlsx = new File(["two"], "second.xlsx");
    fireEvent.drop(zone, { dataTransfer: { files: [docx, xlsx] } });

    await waitFor(() => {
      expect(archiveApi.parseContractCard).toHaveBeenCalledTimes(1);
    });
    expect(archiveApi.parseContractCard).toHaveBeenCalledWith(9, docx);
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
    expect(archiveApi.patchContractCard).not.toHaveBeenCalled();

    const pdf = new File(["three"], "third.pdf", { type: "application/pdf" });
    fireEvent.drop(zone, { dataTransfer: { files: [xlsx, pdf] } });
    await waitFor(() => {
      expect(archiveApi.parseContractCard).toHaveBeenCalledTimes(2);
    });
    expect(archiveApi.parseContractCard).toHaveBeenLastCalledWith(9, xlsx);

    const trailing = new File(["four"], "fourth.docx");
    fireEvent.drop(zone, { dataTransfer: { files: [pdf, trailing] } });
    await waitFor(() => {
      expect(archiveApi.parseContractCard).toHaveBeenCalledTimes(3);
    });
    expect(archiveApi.parseContractCard).toHaveBeenLastCalledWith(9, pdf);
  });

  it.each([
    ["card.docx", ""],
    ["card.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
    ["card.pdf", "application/pdf"],
  ])("shows the file name and source text for %s without zoom", async (name, type) => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { inn: "440117229007" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
      source_text: `Строка ${name}\nИНН 440117229007`,
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], name, { type })] },
    });

    const cardText = await screen.findByLabelText("Текст карточки");
    expect(cardText).toHaveTextContent(`Строка ${name}`);
    expect(cardText).toHaveTextContent("ИНН 440117229007");
    const fileName = screen.getByText(name);
    expect(fileName.compareDocumentPosition(cardText) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Увеличить" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Уменьшить" })).not.toBeInTheDocument();
    expect(screen.queryByRole("img", { name: "Карточка контрагента" })).not.toBeInTheDocument();
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
  });

  it("shows a JPEG card as a picture with zoom", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { inn: "440117229007" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    renderDrawer();

    const file = new File(["jpeg"], "card.jpg", { type: "image/jpeg" });
    fireEvent.change(screen.getByLabelText("Файл карточки"), { target: { files: [file] } });

    const image = await screen.findByRole("img", { name: "Карточка контрагента" });
    const source = screen.getByLabelText("Источник карточки");
    expect(source).toHaveStyle({ position: "sticky", top: "0px" });
    expect(source.parentElement).toHaveStyle({
      gridTemplateColumns: "minmax(320px, 1fr) minmax(0, 1fr)",
    });
    expect(source.compareDocumentPosition(screen.getByLabelText("ИНН")) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(screen.getByText("card.jpg")).toBeInTheDocument();
    expect(image).toHaveStyle({ width: "100%" });
    expect(screen.getByText("100%")).toBeInTheDocument();
    expect(screen.queryByLabelText("Текст карточки")).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Увеличить" }));
    expect(image).toHaveStyle({ width: "125%" });
    expect(screen.getByText("125%")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "По ширине" }));
    expect(image).toHaveStyle({ width: "100%" });
    expect(archiveApi.parseContractCard).toHaveBeenCalledWith(9, file);
  });

  it("changes JPEG zoom on Ctrl+wheel and ignores wheel without Ctrl", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: {},
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Файл карточки"), {
      target: { files: [new File(["jpeg"], "card.jpg", { type: "image/jpeg" })] },
    });

    const image = await screen.findByRole("img", { name: "Карточка контрагента" });
    const scroller = image.parentElement;
    expect(scroller).toBeTruthy();

    fireEvent.wheel(scroller!, { ctrlKey: false, deltaY: -100 });
    expect(image).toHaveStyle({ width: "100%" });

    fireEvent.wheel(scroller!, { ctrlKey: true, deltaY: -1 });
    expect(image).toHaveStyle({ width: "125%" });

    fireEvent.wheel(scroller!, { ctrlKey: true, deltaY: 1 });
    expect(image).toHaveStyle({ width: "100%" });
  });

  it("shows a textless PDF scan as a picture with zoom", async () => {
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { inn: "440117229007" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
      source_text: "",
    });
    renderDrawer();

    fireEvent.change(screen.getByLabelText("Файл карточки"), {
      target: { files: [new File(["scan"], "scan.pdf", { type: "application/pdf" })] },
    });

    const preview = await screen.findByTitle("Карточка контрагента");
    expect(preview.tagName).toBe("IFRAME");
    expect(preview).toHaveStyle({ width: "100%" });
    expect(screen.getByText("scan.pdf")).toBeInTheDocument();
    expect(screen.queryByLabelText("Текст карточки")).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Увеличить" }));
    expect(preview).toHaveStyle({ width: "125%" });
  });

  it("does not erase the INN already in the form when the file is not a counterparty card", async () => {
    vi.mocked(archiveApi.parseContractCard).mockRejectedValue(new Error("это не карточка контрагента"));
    renderDrawer();

    fireEvent.change(screen.getByLabelText("ИНН"), { target: { value: "7707083893" } });
    fireEvent.change(screen.getByLabelText("Файл карточки"), {
      target: { files: [new File(["passport"], "passport.docx")] },
    });

    expect(await screen.findByText("это не карточка контрагента")).toBeInTheDocument();
    expect(screen.getByLabelText("ИНН")).toHaveValue("7707083893");
    expect(screen.getByLabelText("ИНН контрагента")).toHaveValue("440117229007");
    expect(archiveApi.saveContractRequisites).not.toHaveBeenCalled();
    expect(archiveApi.patchContractCard).not.toHaveBeenCalled();
    expect(archiveApi.parseContractCard).toHaveBeenCalledTimes(1);
  });

  it("shows delete under requisites save", () => {
    renderDrawer();

    const save = screen.getByRole("button", { name: "Сохранить реквизиты" });
    const remove = screen.getByRole("button", { name: "Удалить карточку" });
    expect(save.compareDocumentPosition(remove) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it("calls delete only after the exact word УДАЛИТЬ and then closes the card", async () => {
    const onDeleted = vi.fn();
    vi.mocked(archiveApi.deleteContractCard).mockResolvedValue(undefined);
    wrap(<ContractCardDrawer card={card()} onDeleted={onDeleted} />);

    fireEvent.click(screen.getByRole("button", { name: "Удалить карточку" }));
    const confirm = screen.getByRole("button", { name: "Подтвердить удаление" });
    expect(confirm).toBeDisabled();
    fireEvent.click(confirm);
    fireEvent.change(screen.getByLabelText("Подтверждение"), { target: { value: "удалить" } });
    expect(confirm).toBeDisabled();
    fireEvent.click(confirm);
    fireEvent.change(screen.getByLabelText("Подтверждение"), { target: { value: "УДАЛИТЬ " } });
    expect(confirm).toBeDisabled();
    fireEvent.click(confirm);
    expect(archiveApi.deleteContractCard).not.toHaveBeenCalled();
    expect(onDeleted).not.toHaveBeenCalled();

    fireEvent.change(screen.getByLabelText("Подтверждение"), { target: { value: "УДАЛИТЬ" } });
    expect(confirm).toBeEnabled();
    fireEvent.click(confirm);

    await waitFor(() => {
      expect(archiveApi.deleteContractCard).toHaveBeenCalledTimes(1);
    });
    expect(archiveApi.deleteContractCard).toHaveBeenCalledWith(9);
    expect(onDeleted).toHaveBeenCalledTimes(1);
  });
});
