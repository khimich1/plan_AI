import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Import1cDialog } from "@/features/nomenclature/components/Import1cDialog";
import { ApiError } from "@/shared/lib/apiError";
import type { Import1cResponse } from "@/features/nomenclature/types/nomenclature";

const mockImport = vi.fn();

vi.mock("@/features/nomenclature/hooks/useNomenclatureQueries", () => ({
  useImport1cMutation: () => ({
    mutateAsync: mockImport,
    isPending: false,
    isError: false,
    error: null,
    reset: vi.fn(),
  }),
}));

const REPORT: Import1cResponse = {
  product_kind: "pile",
  summary: "+1 GUID, 0 ждут цены, 1 неоднозначных, 0 исчезли из 1С",
  new_guids_count: 1,
  waiting_price: 0,
  ambiguous_count: 1,
  disappeared_count: 0,
  unmatched_1c_count: 1,
  unchanged: 0,
  updated_guids_count: 0,
  list_limit: 50,
  new_guids: [{ product_kind: "pile", mark: "С30.30-3", field: "guid_1c", guid: "aaa-111" }],
  updated_guids: [],
  ambiguous: [
    {
      product_kind: "pile",
      mark: "С40.30-6",
      name: "Сваи С40.30-6",
      guids: ["bbb-222", "ccc-333"],
      note: "дубль в файле",
    },
  ],
  disappeared: [],
  unmatched_1c: [{ name: "Сваи XYZ", guid: "ddd-444", row_index: 12, source_file: "Прайс сваи.xls" }],
};

describe("Import1cDialog", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("imports xls and shows summary plus capped action lists", async () => {
    mockImport.mockResolvedValue(REPORT);
    render(<Import1cDialog open onClose={() => undefined} />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    const file = new File(["xls"], "Прайс сваи.xls", { type: "application/vnd.ms-excel" });
    fireEvent.change(input, { target: { files: [file] } });

    expect(screen.getByText(/Выбран: Прайс сваи\.xls/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Загрузить" }));

    await waitFor(() => {
      expect(mockImport).toHaveBeenCalledWith({ file, productKind: null });
    });

    expect(await screen.findByText(REPORT.summary)).toBeInTheDocument();
    expect(screen.getByText("полная выгрузка")).toBeInTheDocument();
    expect(screen.queryByText(/универсальный отчёт с УИД/)).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Готово" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Загрузить ещё" })).toBeInTheDocument();
    expect(screen.getByText("С30.30-3")).toBeInTheDocument();
    expect(screen.getByText("Сваи С40.30-6 (С40.30-6)")).toBeInTheDocument();
    expect(screen.getByText("Сваи XYZ")).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: /Новые GUID/ })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: /Неоднозначные/ })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: /Нет в прайсе программы/ })).toBeInTheDocument();
  });

  it("sends selected product_kind for mixed ЛМ dumps", async () => {
    mockImport.mockResolvedValue({ ...REPORT, product_kind: "stair_flight", ambiguous: [], unmatched_1c: [], new_guids: [] });
    render(<Import1cDialog open onClose={() => undefined} />);

    fireEvent.change(screen.getByLabelText("Группа изделий"), { target: { value: "stair_flight" } });
    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    const file = new File(["xls"], "Прайс ЛМ.xls", { type: "application/vnd.ms-excel" });
    fireEvent.change(input, { target: { files: [file] } });
    fireEvent.click(screen.getByRole("button", { name: "Загрузить" }));

    await waitFor(() => {
      expect(mockImport).toHaveBeenCalledWith({ file, productKind: "stair_flight" });
    });
  });

  it("hints that a filtered dump is partial and does not mark disappeared", () => {
    render(<Import1cDialog open onClose={() => undefined} />);
    expect(
      screen.getByText(/Выгрузка с отбором — частичный режим: позиции вне файла не помечаем как исчезнувшие/),
    ).toBeInTheDocument();
  });

  it("shows partial mode after a filtered xlsx import", async () => {
    mockImport.mockResolvedValue({
      ...REPORT,
      mode: "partial",
      disappeared_count: 0,
      disappeared: [],
      ambiguous: [],
      unmatched_1c: [],
      new_guids: [],
    });
    render(<Import1cDialog open onClose={() => undefined} />);
    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    fireEvent.change(input, {
      target: {
        files: [
          new File(["x"], "Универсальный отчет.xlsx", {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          }),
        ],
      },
    });
    fireEvent.click(screen.getByRole("button", { name: "Загрузить" }));
    expect(await screen.findByText("частичная выгрузка")).toBeInTheDocument();
  });

  it("accepts universal report xlsx", () => {
    render(<Import1cDialog open onClose={() => undefined} />);
    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    fireEvent.change(input, {
      target: {
        files: [
          new File(["x"], "Универсальный отчет.xlsx", {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          }),
        ],
      },
    });
    expect(screen.getByText(/Выбран: Универсальный отчет\.xlsx/)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Загрузить" })).toBeEnabled();
  });

  it("rejects non-excel locally", () => {
    render(<Import1cDialog open onClose={() => undefined} />);
    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    fireEvent.change(input, {
      target: { files: [new File(["x"], "other.csv", { type: "text/csv" })] },
    });
    expect(screen.getByText(/Нужен файл Excel \(\.xls или \.xlsx\)/)).toBeInTheDocument();
    expect(screen.getByRole("alert")).toHaveTextContent(/Нужен файл Excel \(\.xls или \.xlsx\)/);
    expect(screen.getByRole("button", { name: "Загрузить" })).toBeDisabled();
  });

  it("shows human-readable 422 from the server", async () => {
    mockImport.mockRejectedValue(
      new ApiError("это не выгрузка Прайс-лист 1С: нет листа «Прайс-лист»", 422),
    );
    render(<Import1cDialog open onClose={() => undefined} />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    fireEvent.change(input, {
      target: { files: [new File(["x"], "bad.xls", { type: "application/vnd.ms-excel" })] },
    });
    fireEvent.click(screen.getByRole("button", { name: "Загрузить" }));

    expect(await screen.findByText(/это не выгрузка Прайс-лист 1С: нет листа «Прайс-лист»/)).toBeInTheDocument();
    expect(screen.getByRole("alert")).toHaveTextContent(/это не выгрузка Прайс-лист 1С/);
  });
});
