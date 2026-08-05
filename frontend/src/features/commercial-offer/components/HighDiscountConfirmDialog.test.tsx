import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { HighDiscountConfirmDialog } from "@/features/commercial-offer/components/HighDiscountConfirmDialog";

describe("HighDiscountConfirmDialog", () => {
  it("requires the exact keyword and delegates cancel", () => {
    const onConfirm = vi.fn();
    const onCancel = vi.fn();
    render(
      <HighDiscountConfirmDialog
        open
        discountPercent={16.01}
        isPending={false}
        onConfirm={onConfirm}
        onCancel={onCancel}
      />,
    );

    const confirm = screen.getByRole("button", { name: "OK" });
    expect(confirm).toBeDisabled();
    fireEvent.change(screen.getByRole("textbox"), { target: { value: "подтверждаю" } });
    expect(confirm).toBeDisabled();
    fireEvent.change(screen.getByRole("textbox"), { target: { value: "ПОДТВЕРЖДАЮ" } });
    fireEvent.click(confirm);
    fireEvent.click(screen.getByRole("button", { name: "Отмена" }));
    expect(onConfirm).toHaveBeenCalledOnce();
    expect(onCancel).toHaveBeenCalledOnce();
  });
});
