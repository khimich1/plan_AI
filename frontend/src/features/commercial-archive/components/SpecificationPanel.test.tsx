import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { SpecificationPanel } from "@/features/commercial-archive/components/SpecificationPanel";
import type { SpecificationView } from "@/features/commercial-archive/types/archive";

const PERCENT_PARAGRAPH =
  "500,00 (пятьсот) рублей — предварительная оплата в размере 50% до 01.10.2026 включительно.";

function percentView(overrides: Partial<SpecificationView> = {}): SpecificationView {
  return {
    saved: true,
    stale_custom: false,
    has_piles: false,
    concrete_grade: null,
    payment_paragraph: PERCENT_PARAGRAPH,
    term_paragraph: "Поставщик обязуется поставить товар не позднее 1 октября 2026 года.",
    delivery_paragraph: "Самовывоз со склада готовой продукции.",
    spec_date: "2026-09-18",
    choice: {
      payment: "split_50_50",
      term: "by_date",
      delivery: "pickup",
      payment_date: "2026-10-01",
      payment_days: 3,
      term_date: "2026-10-15",
    },
    ...overrides,
  };
}

afterEach(() => {
  cleanup();
});

describe("SpecificationPanel", () => {
  it("shows a percent paragraph as text, not a textarea", () => {
    render(<SpecificationPanel view={percentView()} busy={false} error={null} onSave={vi.fn()} onDownload={vi.fn()}
        onDownloadPdf={vi.fn()} />);

    expect(screen.getByText(PERCENT_PARAGRAPH).tagName).toBe("P");
    expect(screen.queryByRole("textbox", { name: "Свой текст" })).not.toBeInTheDocument();
    expect(document.querySelector("textarea")).toBeNull();
  });

  it("uses a textarea only for custom payment text", () => {
    render(
      <SpecificationPanel
        view={percentView({
          payment_paragraph: "Оплата по графику заказчика.",
          choice: {
            payment: "custom",
            term: "by_date",
            delivery: "pickup",
            custom_text: "Оплата по графику заказчика.",
            term_date: "2026-10-15",
          },
        })}
        busy={false}
        error={null}
        onSave={vi.fn()}
        onDownload={vi.fn()}
        onDownloadPdf={vi.fn()}
      />,
    );

    expect(screen.getByRole("textbox", { name: "Свой текст" })).toHaveValue("Оплата по графику заказчика.");
  });

  it("hides pile rhythm and concrete grade when the server says there are no piles", () => {
    render(<SpecificationPanel view={percentView()} busy={false} error={null} onSave={vi.fn()} onDownload={vi.fn()}
        onDownloadPdf={vi.fn()} />);

    expect(screen.queryByRole("button", { name: "По N свай" })).not.toBeInTheDocument();
    expect(screen.queryByText(/Марка бетона/)).not.toBeInTheDocument();
  });

  it("shows pile rhythm and concrete grade only when the server sends them", () => {
    render(
      <SpecificationPanel
        view={percentView({ has_piles: true, concrete_grade: "В25" })}
        busy={false}
        error={null}
        onSave={vi.fn()}
        onDownload={vi.fn()}
        onDownloadPdf={vi.fn()}
      />,
    );

    expect(screen.getByRole("button", { name: "По N свай" })).toBeInTheDocument();
    expect(screen.getByText("Марка бетона: В25")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "в неделю" })).not.toBeInTheDocument();
  });

  it("shows save, excel and pdf, and pdf submits the same draft as excel", async () => {
    const onDownload = vi.fn().mockResolvedValue(undefined);
    const onDownloadPdf = vi.fn().mockResolvedValue(undefined);
    render(
      <SpecificationPanel
        view={percentView()}
        busy={false}
        error={null}
        onSave={vi.fn()}
        onDownload={onDownload}
        onDownloadPdf={onDownloadPdf}
      />,
    );

    const save = screen.getByRole("button", { name: "Сохранить" });
    const excel = screen.getByRole("button", { name: "Скачать Excel" });
    const pdf = screen.getByRole("button", { name: "Скачать PDF" });
    expect(save.compareDocumentPosition(excel) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(excel.compareDocumentPosition(pdf) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();

    fireEvent.click(excel);
    await waitFor(() => {
      expect(onDownload).toHaveBeenCalledTimes(1);
    });
    fireEvent.click(pdf);
    await waitFor(() => {
      expect(onDownloadPdf).toHaveBeenCalledTimes(1);
    });
    expect(onDownloadPdf).toHaveBeenCalledWith(onDownload.mock.calls[0][0]);
  });
});
