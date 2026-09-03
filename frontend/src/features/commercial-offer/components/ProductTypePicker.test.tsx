import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import type { ComponentProps } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ProductTypePicker } from "@/features/commercial-offer/components/ProductTypePicker";
import type { ProductType } from "@/features/commercial-offer/types/commercialOffer";

afterEach(() => {
  cleanup();
});

describe("ProductTypePicker", () => {
  describe("create mode (default)", () => {
    it("renders plates, piles, steps, marches, bridge piles, and fbs options", () => {
      render(<ProductTypePicker onSelect={vi.fn()} />);

      expect(screen.getByRole("button", { name: /Плиты/i })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: /цельные железобетонные сваи/i })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: /Ступени/i })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: /лестничные марши/i })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: /мостовые железобетонные сваи/i })).toBeInTheDocument();
      expect(screen.getByRole("button", { name: /фундаментные блоки ФБС/i })).toBeInTheDocument();
    });

    it("calls onSelect with steps when the steps card is clicked", () => {
      const onSelect = vi.fn();

      render(<ProductTypePicker onSelect={onSelect} />);
      const stepsButtons = screen.getAllByRole("button", { name: /Ступени/i });
      fireEvent.click(stepsButtons[stepsButtons.length - 1]!);

      expect(onSelect).toHaveBeenCalledWith("steps");
    });

    it("calls onSelect with marches when the marches card is clicked", () => {
      const onSelect = vi.fn();

      render(<ProductTypePicker onSelect={onSelect} />);
      const marchButtons = screen.getAllByRole("button", { name: /лестничные марши/i });
      fireEvent.click(marchButtons[marchButtons.length - 1]!);

      expect(onSelect).toHaveBeenCalledWith("marches");
    });

    it("calls onSelect with bridge_piles when the bridge piles card is clicked", () => {
      const onSelect = vi.fn();

      render(<ProductTypePicker onSelect={onSelect} />);
      fireEvent.click(screen.getByRole("button", { name: /мостовые железобетонные сваи/i }));

      expect(onSelect).toHaveBeenCalledWith("bridge_piles");
    });

    it("calls onSelect with fbs when the FBS card is clicked", () => {
      const onSelect = vi.fn();

      render(<ProductTypePicker onSelect={onSelect} />);
      fireEvent.click(screen.getByRole("button", { name: /фундаментные блоки ФБС/i }));

      expect(onSelect).toHaveBeenCalledWith("fbs");
    });

    it("shows create heading and neutral subtitle without single-type restriction copy", () => {
      render(<ProductTypePicker onSelect={vi.fn()} />);

      expect(screen.getByRole("heading", { name: /Создание коммерческого предложения/i })).toBeInTheDocument();
      expect(screen.getByText(/Выберите тип продукции/i)).toBeInTheDocument();
      expect(screen.queryByText(/в одном КП только один тип/i)).not.toBeInTheDocument();
    });

    it("does not show append strip, checkmarks, or action buttons", () => {
      render(<ProductTypePicker onSelect={vi.fn()} />);

      expect(screen.queryByText(/Уже в КП/i)).not.toBeInTheDocument();
      expect(screen.queryByRole("button", { name: /К результату/i })).not.toBeInTheDocument();
      expect(screen.queryByRole("button", { name: /Добавить .* в КП/i })).not.toBeInTheDocument();
      expect(screen.queryByRole("button", { name: /Показать позиции/i })).not.toBeInTheDocument();
      expect(screen.queryByLabelText(/Уже добавлен/i)).not.toBeInTheDocument();
    });
  });

  describe("append mode", () => {
    const selectedTypes: ProductType[] = ["plates", "piles"];
    const orderLines = [
      { product_type: "plates", name: "ПБ 60-12-8", qty: 10 },
      { product_type: "plates", mark: "ПБ 58-15-8", qty: 4 },
      { product_type: "piles", mark: "С80.30-8", qty: 2 },
      { product_type: "fbs", name: "ФБС 24-4-6", qty: 1 },
    ];

    const renderAppend = (overrides: Partial<ComponentProps<typeof ProductTypePicker>> = {}) => {
      const onSelect = vi.fn();
      const onBackToResult = vi.fn();
      const result = render(
        <ProductTypePicker
          mode="append"
          selectedProductTypes={selectedTypes}
          orderLines={orderLines}
          managerName="Иванов И.И."
          clientName="ООО Ромашка"
          onSelect={onSelect}
          onBackToResult={onBackToResult}
          {...overrides}
        />,
      );
      return { ...result, onSelect, onBackToResult };
    };

    it("shows manager and client in header instead of create heading", () => {
      renderAppend();

      expect(screen.getByRole("heading", { name: /Иванов И\.И\./i })).toBeInTheDocument();
      expect(screen.getByText(/ООО Ромашка/i)).toBeInTheDocument();
      expect(screen.queryByRole("heading", { name: /Создание коммерческого предложения/i })).not.toBeInTheDocument();
      expect(screen.getByText(/Выберите тип продукции для дополнения текущего КП/i)).toBeInTheDocument();
    });

    it("shows «Уже в КП» strip with Russian labels of selected types", () => {
      renderAppend();

      expect(screen.getByText(/Уже в КП:\s*Плиты\s*·\s*Сваи/i)).toBeInTheDocument();
    });

    it("calls onBackToResult when «К результату» is clicked", () => {
      const { onBackToResult } = renderAppend();

      fireEvent.click(screen.getByRole("button", { name: /К результату/i }));

      expect(onBackToResult).toHaveBeenCalledTimes(1);
    });

    it("selects an unselected type via whole-tile click (S2)", () => {
      const { onSelect } = renderAppend();

      fireEvent.click(screen.getByRole("button", { name: /фундаментные блоки ФБС/i }));

      expect(onSelect).toHaveBeenCalledWith("fbs");
    });

    it("does not call onSelect when clicking the background of an already-selected tile (S3)", () => {
      const { onSelect, container } = renderAppend();

      const platesTile = container.querySelector('[data-product-type="plates"]');
      expect(platesTile).toBeTruthy();
      fireEvent.click(platesTile!);

      expect(onSelect).not.toHaveBeenCalled();
    });

    it("calls onSelect when (+) is clicked on a selected type (S3)", () => {
      const { onSelect } = renderAppend();

      fireEvent.click(screen.getByRole("button", { name: /Добавить плиты в КП/i }));

      expect(onSelect).toHaveBeenCalledWith("plates");
    });

    it("opens Drawer with name/mark and qty for the type when (i) is clicked (S4)", () => {
      renderAppend();

      fireEvent.click(screen.getByRole("button", { name: /Показать позиции: плиты/i }));

      const dialog = screen.getByRole("dialog");
      expect(dialog).toBeInTheDocument();
      expect(within(dialog).getByText("ПБ 60-12-8")).toBeInTheDocument();
      expect(within(dialog).getByText("ПБ 58-15-8")).toBeInTheDocument();
      expect(within(dialog).getByText("10")).toBeInTheDocument();
      expect(within(dialog).getByText("4")).toBeInTheDocument();
      expect(within(dialog).queryByText("С80.30-8")).not.toBeInTheDocument();
      expect(within(dialog).queryByText(/₽|цена|сумма/i)).not.toBeInTheDocument();
    });

    it("opens Drawer on the left with numbered rows (№, name, qty)", () => {
      renderAppend();

      fireEvent.click(screen.getByRole("button", { name: /Показать позиции: плиты/i }));

      const dialog = screen.getByRole("dialog");
      expect(dialog).toHaveClass("app-drawer--left");
      expect(within(dialog).getByRole("columnheader", { name: "№" })).toBeInTheDocument();
      expect(within(dialog).getByRole("columnheader", { name: "Наименование" })).toBeInTheDocument();
      expect(within(dialog).getByRole("columnheader", { name: "Кол-во" })).toBeInTheDocument();

      const rows = within(dialog).getAllByRole("row");
      // header + 2 plate lines
      expect(rows).toHaveLength(3);
      expect(within(rows[1]!).getByText("1")).toBeInTheDocument();
      expect(within(rows[1]!).getByText("ПБ 60-12-8")).toBeInTheDocument();
      expect(within(rows[2]!).getByText("2")).toBeInTheDocument();
      expect(within(rows[2]!).getByText("ПБ 58-15-8")).toBeInTheDocument();
    });

    it("shows checkmark indicator on selected tiles", () => {
      renderAppend();

      expect(screen.getByLabelText(/Уже добавлен: плиты/i)).toBeInTheDocument();
      expect(screen.getByLabelText(/Уже добавлен: сваи/i)).toBeInTheDocument();
      expect(screen.queryByLabelText(/Уже добавлен: фбс/i)).not.toBeInTheDocument();
    });
  });
});
