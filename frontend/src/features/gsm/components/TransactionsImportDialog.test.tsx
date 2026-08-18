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

    expect(await screen.findByText(/расхождений по файлам: 1/)).toBeInTheDocument();
    const badRow = screen.getByText("bad.xls").closest("tr") as HTMLElement;
    expect(badRow).toHaveAttribute("data-mismatch", "true");
    expect(within(badRow).getByText(/Расхождение итогов/)).toBeInTheDocument();
    expect(within(badRow).getByText(/Неизвестные карты: 9999/)).toBeInTheDocument();

    const okRow = screen.getByText("ok.xls").closest("tr") as HTMLElement;
    expect(okRow).toHaveAttribute("data-mismatch", "false");
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
