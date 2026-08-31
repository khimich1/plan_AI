import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { TransactionsImportDialog } from "@/features/gsm/components/TransactionsImportDialog";
import { ApiError } from "@/shared/lib/apiError";
import type { TransactionImportReport } from "@/features/gsm/types/gsm";

const mockImport = vi.fn();

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useImportGsmTransactionsMutation: () => ({
    mutateAsync: mockImport,
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
}));

const REPORT: TransactionImportReport = {
  rows_inserted: 12,
  rows_duplicate: 1,
  files: [
    {
      filename: "ok.xls",
      rows_total: 10,
      rows_inserted: 10,
      rows_duplicate: 0,
      sum_liters: 100,
      sum_amount: 5000,
      footer_liters: 100,
      footer_amount: 5000,
      warnings: [],
      unmatched_cards: [],
    },
    {
      filename: "bad.xls",
      rows_total: 5,
      rows_inserted: 2,
      rows_duplicate: 1,
      sum_liters: 40,
      sum_amount: 2000,
      footer_liters: 45,
      footer_amount: 2000,
      warnings: ["bad.xls: Кол-во 40.00 ≠ Итоги 45.00"],
      unmatched_cards: ["9999"],
    },
  ],
};

const ALL_DUPLICATE: TransactionImportReport = {
  rows_inserted: 0,
  rows_duplicate: 3,
  files: [
    {
      filename: "transactions_excel.xls",
      rows_total: 3,
      rows_inserted: 0,
      rows_duplicate: 3,
      sum_liters: 115,
      sum_amount: 9199.98,
      footer_liters: 115,
      footer_amount: 9199.98,
      warnings: [],
      unmatched_cards: [],
    },
  ],
};

describe("TransactionsImportDialog", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("imports multiple files and highlights mismatch rows", async () => {
    mockImport.mockResolvedValue(REPORT);
    const onImported = vi.fn();
    render(<TransactionsImportDialog open onClose={() => undefined} onImported={onImported} />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    const fileA = new File(["a"], "ok.xls", { type: "application/vnd.ms-excel" });
    const fileB = new File(["b"], "bad.xls", { type: "application/vnd.ms-excel" });
    fireEvent.change(input, { target: { files: [fileA, fileB] } });

    expect(screen.getByText(/Выбрано:.*ok\.xls/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: /Импортировать \(2\)/ }));

    await waitFor(() => {
      expect(mockImport).toHaveBeenCalledWith([fileA, fileB]);
    });
    expect(onImported).toHaveBeenCalledWith(REPORT);

    expect(await screen.findByText(/Расхождение итогов по 1 файлам/)).toBeInTheDocument();
    expect(screen.getByText(/Добавлено 12 операций/)).toBeInTheDocument();
    expect(screen.queryByText(/дубл|вставлен/i)).not.toBeInTheDocument();

    expect(screen.getByRole("columnheader", { name: "Прочитано" })).toBeInTheDocument();
    expect(screen.getByRole("columnheader", { name: "Добавлено" })).toBeInTheDocument();
    expect(screen.getByRole("columnheader", { name: "Уже были" })).toBeInTheDocument();

    const badRow = screen.getByText("bad.xls").closest("tr") as HTMLElement;
    expect(badRow).toHaveAttribute("data-mismatch", "true");
    expect(within(badRow).getByText(/Расхождение итогов/)).toBeInTheDocument();
    expect(within(badRow).getByText(/Неизвестные карты: 9999/)).toBeInTheDocument();
    expect(within(badRow).getByText("5")).toBeInTheDocument();
    expect(within(badRow).getByText("2")).toBeInTheDocument();
    expect(within(badRow).getByText("1")).toBeInTheDocument();

    const okRow = screen.getByText("ok.xls").closest("tr") as HTMLElement;
    expect(okRow).toHaveAttribute("data-mismatch", "false");
  });

  it("explains all-duplicate import without jargon", async () => {
    mockImport.mockResolvedValue(ALL_DUPLICATE);
    render(<TransactionsImportDialog open onClose={() => undefined} />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    fireEvent.change(input, {
      target: {
        files: [new File(["x"], "transactions_excel.xls", { type: "application/vnd.ms-excel" })],
      },
    });
    fireEvent.click(screen.getByRole("button", { name: /Импортировать \(1\)/ }));

    expect(
      await screen.findByText(
        /Новых операций нет: все 3 уже есть в журнале\. Повторная загрузка того же файла ничего не меняет\./,
      ),
    ).toBeInTheDocument();
    expect(screen.queryByText(/дубл|вставлен/i)).not.toBeInTheDocument();

    const row = screen.getByText("transactions_excel.xls").closest("tr") as HTMLElement;
    const cells = within(row).getAllByRole("cell");
    // Файл, Прочитано=3, Добавлено=0, Уже были=3, …
    expect(cells[1]).toHaveTextContent("3");
    expect(cells[2]).toHaveTextContent("0");
    expect(cells[3]).toHaveTextContent("3");
  });

  it("shows safe-reload hint before import", () => {
    render(<TransactionsImportDialog open onClose={() => undefined} />);
    expect(screen.getByText(/Повторная загрузка безопасна/i)).toBeInTheDocument();
  });

  it("shows human-readable API error on failed import", async () => {
    mockImport.mockRejectedValue(
      new ApiError("dup", 422, "card_number «7001» already exists", "gsm_card_duplicate"),
    );

    render(<TransactionsImportDialog open onClose={() => undefined} />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    fireEvent.change(input, {
      target: { files: [new File(["x"], "tx.xls", { type: "application/vnd.ms-excel" })] },
    });
    fireEvent.click(screen.getByRole("button", { name: /Импортировать \(1\)/ }));

    expect(await screen.findByText(/Карта «7001» уже существует/)).toBeInTheDocument();
  });
});
