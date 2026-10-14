import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { TimesheetImportDialog } from "@/features/gsm/components/TimesheetImportDialog";
import { ApiError } from "@/shared/lib/apiError";
import type { TimesheetImportReport } from "@/features/gsm/types/gsm";

const mockImport = vi.fn();

vi.mock("@/features/gsm/hooks/useGsmQueries", () => ({
  useImportTimesheetMutation: () => ({
    mutateAsync: mockImport,
    isPending: false,
    reset: vi.fn(),
  }),
}));

const REPORT: TimesheetImportReport = {
  period: "2026-07",
  rows_total: 7,
  matched: 5,
  unmatched: [
    { full_name: "Офисная Сотрудница", personnel_number: "999" },
    { full_name: "Без Табельного", personnel_number: null },
  ],
  codes_summary: { Я: 100, ОТ: 9 },
};

describe("TimesheetImportDialog", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("shows five matched drivers and office rows without a driver", async () => {
    mockImport.mockResolvedValue(REPORT);
    render(<TimesheetImportDialog open onClose={() => undefined} />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    const file = new File([new Uint8Array([1])], "табель от 31.07.2026.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    fireEvent.change(input, { target: { files: [file] } });
    fireEvent.click(screen.getByRole("button", { name: "Загрузить" }));

    expect(await screen.findByText(/сопоставлено 5 из 7/)).toBeInTheDocument();
    expect(screen.getByText("Офисная Сотрудница")).toBeInTheDocument();
    expect(screen.getByText("Без Табельного")).toBeInTheDocument();
    expect(screen.queryByLabelText("Месяц табеля")).not.toBeInTheDocument();
  });

  it("asks for a month when the filename has no date and shows the 400 text", async () => {
    mockImport.mockRejectedValue(
      new ApiError(
        "month",
        400,
        "Укажите месяц табеля: в имени файла нет даты «от ДД.ММ.ГГГГ».",
        "gsm_timesheet_month_required",
      ),
    );
    render(<TimesheetImportDialog open onClose={() => undefined} />);

    const input = document.querySelector('input[type="file"]') as HTMLInputElement;
    const file = new File([new Uint8Array([1])], "табель.xlsx", {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    fireEvent.change(input, { target: { files: [file] } });

    expect(screen.getByLabelText("Месяц табеля")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Загрузить" }));

    await waitFor(() => {
      expect(mockImport).toHaveBeenCalledWith({ file, month: undefined });
    });
    expect(
      await screen.findByRole("alert"),
    ).toHaveTextContent("Укажите месяц табеля: в имени файла нет даты «от ДД.ММ.ГГГГ».");
  });
});