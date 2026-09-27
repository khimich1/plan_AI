import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import {
  SupplyContractArchiveButton,
  SupplyContractDrawer,
} from "@/features/commercial-archive/components/SupplyContractDrawer";
import type {
  SupplyContract,
  SupplyContractParseResult,
} from "@/features/commercial-archive/types/supplyContract";

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    getSupplyContract: vi.fn(),
    createSupplyContract: vi.fn(),
    parseSupplyContract: vi.fn(),
    downloadSupplyContract: vi.fn(),
    listSupplyContracts: vi.fn(),
    patchSupplyContract: vi.fn(),
    replaceSupplyContract: vi.fn(),
  },
}));

const contract: SupplyContract = {
  id: 4,
  counterparty_id: 15,
  number: "1028/09/26",
  contract_date: "2026-09-25",
  manager_name: "Иван Иванов",
  status: "нет",
  legal_form: "ooo",
  full_name: "Ромашка",
  short_name: "Ромашка",
  signatory_name: "Иванов Иван",
  email: "a@b.ru",
};

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("SupplyContractArchiveButton", () => {
  it("disables the contract button without a counterparty", () => {
    render(<SupplyContractArchiveButton counterpartyId={null} onClick={vi.fn()} />);

    const button = screen.getByRole("button", { name: "Договор" });
    expect(button).toBeDisabled();
    expect(button).toHaveAttribute("title", "Сначала занесите контрагента из 1С");
  });

  it("enables the contract button when a counterparty is set", () => {
    render(<SupplyContractArchiveButton counterpartyId={15} onClick={vi.fn()} />);
    expect(screen.getByRole("button", { name: "Договор" })).toBeEnabled();
  });
});

describe("SupplyContractDrawer", () => {
  it("shows the number and no create form when a contract exists", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(contract);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);

    wrap(
      <SupplyContractDrawer
        open
        onClose={vi.fn()}
        kpId={42}
        counterpartyId={15}
        customerName="ООО Тест"
      />,
    );

    expect(await screen.findByText("Номер: 1028/09/26")).toBeInTheDocument();
    expect(screen.getByText("Дата: 2026-09-25")).toBeInTheDocument();
    expect(screen.getByText("Статус: нет")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Скачать docx" })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Подтвердить" })).not.toBeInTheDocument();
    expect(screen.queryByLabelText("Файл карточки")).not.toBeInTheDocument();
  });

  it("confirms an empty contract with POST and shows the issued number", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.createSupplyContract).mockResolvedValue(contract);

    wrap(
      <SupplyContractDrawer
        open
        onClose={vi.fn()}
        kpId={42}
        counterpartyId={15}
        customerName="Ромашка"
        customerInn="760401001"
        customerKpp="760401001"
      />,
    );

    expect(await screen.findByRole("button", { name: "Подтвердить" })).toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Полное наименование"), { target: { value: "Ромашка" } });
    fireEvent.change(screen.getByLabelText("ФИО подписанта"), { target: { value: "Иванов Иван" } });
    fireEvent.change(screen.getByLabelText("Должность"), { target: { value: "Директора" } });
    fireEvent.change(screen.getByLabelText("ОГРН"), { target: { value: "1027600000001" } });
    fireEvent.change(screen.getByLabelText("Юридический адрес"), { target: { value: "Ярославль" } });
    fireEvent.change(screen.getByLabelText("E-mail"), { target: { value: "a@b.ru" } });
    fireEvent.change(screen.getByLabelText("Банк"), { target: { value: "Промсвязьбанк" } });
    fireEvent.change(screen.getByLabelText("Расчётный счёт"), { target: { value: "40702810000000000001" } });
    fireEvent.change(screen.getByLabelText("Корсчёт"), { target: { value: "30101810000000000760" } });
    fireEvent.change(screen.getByLabelText("БИК"), { target: { value: "044525555" } });
    fireEvent.click(screen.getByRole("button", { name: "Подтвердить" }));

    await waitFor(() => {
      expect(archiveApi.createSupplyContract).toHaveBeenCalledWith(
        42,
        expect.objectContaining({ full_name: "Ромашка", inn: "760401001", short_name: "Ромашка" }),
      );
    });
    expect(await screen.findByText("Номер: 1028/09/26")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Подтвердить" })).not.toBeInTheDocument();
  });

  it("keeps the source file beside the fields and highlights doubtful ones", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "760401001", email: "a@b.ru" },
      accounts: [],
      doubtful: ["inn"],
      verify_failed: false,
    });

    wrap(
      <SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />,
    );

    const input = await screen.findByLabelText("Файл карточки");
    const file = new File(["card"], "card.png", { type: "image/png" });
    fireEvent.change(input, { target: { files: [file] } });

    expect(await screen.findByRole("img", { name: "Карточка контрагента" })).toBeInTheDocument();
    expect(screen.getByText("card.png")).toBeInTheDocument();
    expect(screen.getByLabelText("ИНН").closest("[data-doubtful='true']")).toBeTruthy();
    expect(screen.getByLabelText("Полное наименование")).toBeInTheDocument();
  });

  it("lists two accounts and writes only the chosen one into the empty field", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "7604010011", account: "" },
      accounts: ["40702810000000000007", "40802810000000000015"],
      doubtful: ["account"],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    expect(await screen.findByRole("button", { name: "40702810000000000007" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "40802810000000000015" })).toBeInTheDocument();
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("");

    fireEvent.click(screen.getByRole("button", { name: "40802810000000000015" }));
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("40802810000000000015");
  });

  it("drops the first file onto the contract screen and ignores the second", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "7604010011" },
      accounts: [],
      doubtful: [],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    const zone = await screen.findByLabelText("Зона файла карточки");
    const first = new File(["one"], "first.docx");
    const second = new File(["two"], "second.docx");
    fireEvent.drop(zone, { dataTransfer: { files: [first, second] } });

    await waitFor(() => {
      expect(archiveApi.parseSupplyContract).toHaveBeenCalledTimes(1);
    });
    expect(archiveApi.parseSupplyContract).toHaveBeenCalledWith(42, first);
  });

  it("does not fill the form when parsing fails", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockRejectedValue(new Error("это не карточка контрагента"));

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["passport"], "passport.docx")] },
    });

    expect(await screen.findByText("это не карточка контрагента")).toBeInTheDocument();
    expect(screen.getByLabelText("ИНН")).toHaveValue("");
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("");
  });

  it("warns only about the empty account and leaves filled names without a doubtful frame", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: {
        full_name: "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ ДСТРОЙ 76",
        short_name: "ДСТРОЙ 76 ООО",
        account: "",
      },
      accounts: [],
      doubtful: ["full_name", "short_name", "account"],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    expect(await screen.findByText("Проверьте по карточке: расчётный счёт.")).toBeInTheDocument();
    expect(screen.getByLabelText("Полное наименование").closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("Краткое наименование").closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("Расчётный счёт")).toHaveValue("");
    expect(screen.getByLabelText("Расчётный счёт").closest("[data-doubtful='true']")).toBeTruthy();
  });

  it("keeps the charter option when parsing returns the genitive wording", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { authority_basis: "Устава" },
      accounts: [],
      doubtful: [],
      verify_failed: false,
    } as SupplyContractParseResult);

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    await waitFor(() => {
      expect(screen.getByLabelText("Основание")).toHaveValue("устав");
    });
  });

  it("keeps the counterparty short name when parsing returns an empty short_name", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { short_name: "", full_name: "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ ДСТРОЙ 76" },
      accounts: [],
      doubtful: ["short_name", "full_name"],
      verify_failed: false,
    });

    wrap(
      <SupplyContractDrawer
        open
        onClose={vi.fn()}
        kpId={42}
        counterpartyId={15}
        customerName="ДСТРОЙ 76 ООО"
      />,
    );

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    await waitFor(() => {
      expect(screen.getByLabelText("Полное наименование")).toHaveValue(
        "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ ДСТРОЙ 76",
      );
    });
    expect(screen.getByLabelText("Краткое наименование")).toHaveValue("ДСТРОЙ 76 ООО");
    expect(screen.getByLabelText("Краткое наименование").closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("Полное наименование").closest("[data-doubtful='true']")).toBeNull();
  });

  it("does not name the settlement account when two accounts are listed", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "7604010011", bik: "" },
      accounts: ["40702810000000000007", "40802810000000000015"],
      doubtful: ["account", "bik"],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    expect(await screen.findByText("Проверьте по карточке: БИК.")).toBeInTheDocument();
  });

  it("lists doubtful digits in card order", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: {},
      accounts: [],
      doubtful: ["bik", "corr_account", "account", "ogrn", "kpp", "inn"],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    expect(
      await screen.findByText("Проверьте по карточке: ИНН, КПП, ОГРН, расчётный счёт, корсчёт, БИК."),
    ).toBeInTheDocument();
  });

  it("does not warn when none of the six digits are doubtful", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "7604010011", full_name: "Ромашка" },
      accounts: [],
      doubtful: ["full_name", "email"],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    await waitFor(() => {
      expect(screen.getByLabelText("ИНН")).toHaveValue("7604010011");
    });
    expect(screen.queryByText(/Проверьте по карточке/)).not.toBeInTheDocument();
    expect(screen.getByLabelText("Полное наименование").closest("[data-doubtful='true']")).toBeNull();
  });

  it("zooms the card photo above 100% and fits it back to width", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "7604010011" },
      accounts: [],
      doubtful: [],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
    });

    const image = await screen.findByRole("img", { name: "Карточка контрагента" });
    const source = screen.getByLabelText("Источник карточки");
    expect(source).toHaveStyle({ position: "sticky", top: "0px" });
    expect(source.parentElement).toHaveStyle({
      gridTemplateColumns: "minmax(320px, 1fr) minmax(0, 1fr)",
    });
    expect(image).toHaveStyle({ width: "100%" });
    expect(screen.getByText("100%")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Увеличить" }));
    expect(image).toHaveStyle({ width: "125%" });
    expect(screen.getByText("125%")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "По ширине" }));
    expect(image).toHaveStyle({ width: "100%" });
    expect(screen.getByText("100%")).toBeInTheDocument();
  });

  it("changes zoom on Ctrl+wheel and ignores wheel without Ctrl", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: {},
      accounts: [],
      doubtful: [],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.png", { type: "image/png" })] },
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

  it("resets zoom to 100% when another photo is loaded", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: {},
      accounts: [],
      doubtful: [],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    const input = await screen.findByLabelText("Файл карточки");
    fireEvent.change(input, {
      target: { files: [new File(["one"], "one.png", { type: "image/png" })] },
    });

    const image = await screen.findByRole("img", { name: "Карточка контрагента" });
    fireEvent.click(screen.getByRole("button", { name: "Увеличить" }));
    expect(image).toHaveStyle({ width: "125%" });

    fireEvent.change(input, {
      target: { files: [new File(["two"], "two.png", { type: "image/png" })] },
    });

    await waitFor(() => {
      expect(screen.getByText("two.png")).toBeInTheDocument();
    });
    expect(screen.getByRole("img", { name: "Карточка контрагента" })).toHaveStyle({ width: "100%" });
    expect(screen.getByText("100%")).toBeInTheDocument();
  });

  it("does not show zoom buttons for a docx card", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "7604010011" },
      accounts: [],
      doubtful: [],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    await waitFor(() => {
      expect(screen.getByText("card.docx")).toBeInTheDocument();
    });
    expect(screen.queryByRole("button", { name: "Увеличить" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Уменьшить" })).not.toBeInTheDocument();
    expect(screen.queryByRole("img", { name: "Карточка контрагента" })).not.toBeInTheDocument();
  });

  it("shows source text for a docx card and no zoom buttons", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "7604010011" },
      accounts: [],
      doubtful: [],
      verify_failed: false,
      source_text: "Строка карточки\nИНН 7604010011",
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    const cardText = await screen.findByLabelText("Текст карточки");
    expect(cardText).toHaveTextContent("Строка карточки");
    expect(cardText).toHaveTextContent("ИНН 7604010011");
    const fileName = screen.getByText("card.docx");
    expect(fileName.compareDocumentPosition(cardText) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(screen.queryByRole("button", { name: "Увеличить" })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Уменьшить" })).not.toBeInTheDocument();
  });

  it("names the address and signatory and drops bank fields when two bundles are listed", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
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

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    expect(await screen.findByText("Проверьте по карточке: юридический адрес, ФИО.")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Альфа 40702810000000000007" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "ВТБ 40702810900000000009" })).toBeInTheDocument();
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
    expect(screen.getByLabelText("Расчётный счёт").closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("Корсчёт").closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("БИК").closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("Юридический адрес").closest("[data-doubtful='true']")).toBeTruthy();
    expect(screen.getByLabelText("ФИО подписанта").closest("[data-doubtful='true']")).toBeTruthy();
    expect(screen.getByRole("button", { name: "Подтвердить" })).toBeEnabled();
  });

  it("clears the address frame after an edit and keeps it after focus and tab", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { legal_address: "ул. Ленина, 1", signatory_name: "Иванов Иван Иванович" },
      accounts: [],
      doubtful: ["legal_address", "signatory_name"],
      verify_failed: false,
    });

    wrap(<SupplyContractDrawer open onClose={vi.fn()} kpId={42} counterpartyId={15} />);

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    const address = await screen.findByLabelText("Юридический адрес");
    expect(address.closest("[data-doubtful='true']")).toBeTruthy();
    expect(screen.getByText("Проверьте по карточке: юридический адрес, ФИО.")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Подтвердить" })).toBeEnabled();

    fireEvent.focus(address);
    fireEvent.keyDown(address, { key: "Tab" });
    expect(address.closest("[data-doubtful='true']")).toBeTruthy();

    fireEvent.change(address, { target: { value: "ул. Ленина, 2" } });
    expect(address.closest("[data-doubtful='true']")).toBeNull();
    expect(screen.getByLabelText("ФИО подписанта").closest("[data-doubtful='true']")).toBeTruthy();
  });

  it("keeps the counterparty inn yellow when the card does not confirm it", async () => {
    vi.mocked(archiveApi.getSupplyContract).mockResolvedValue(null);
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([]);
    vi.mocked(archiveApi.parseSupplyContract).mockResolvedValue({
      fields: { inn: "" },
      accounts: [],
      doubtful: ["inn"],
      verify_failed: false,
    });

    wrap(
      <SupplyContractDrawer
        open
        onClose={vi.fn()}
        kpId={42}
        counterpartyId={15}
        customerInn="7707083893"
      />,
    );

    fireEvent.change(await screen.findByLabelText("Файл карточки"), {
      target: { files: [new File(["card"], "card.docx")] },
    });

    const inn = await screen.findByLabelText("ИНН");
    expect(inn).toHaveValue("7707083893");
    expect(inn.closest("[data-doubtful='true']")).toBeTruthy();
  });
});
