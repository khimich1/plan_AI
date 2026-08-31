import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { GsmTabs } from "@/features/gsm/components/GsmTabs";
import type { GsmTab } from "@/features/gsm/types/gsm";

describe("GsmTabs", () => {
  afterEach(() => {
    cleanup();
  });

  it("marks the active tab and calls onChange", () => {
    const onChange = vi.fn();
    let value: GsmTab = "overview";

    const { rerender } = render(<GsmTabs value={value} onChange={onChange} />);

    expect(screen.getByRole("tab", { name: "Обзор" })).toHaveAttribute("aria-selected", "true");
    expect(screen.getByRole("tab", { name: "Транзакции" })).toHaveAttribute("aria-selected", "false");

    fireEvent.click(screen.getByRole("tab", { name: "Справочники" }));
    expect(onChange).toHaveBeenCalledWith("registries");

    value = "registries";
    rerender(<GsmTabs value={value} onChange={onChange} />);
    expect(screen.getByRole("tab", { name: "Справочники" })).toHaveAttribute("aria-selected", "true");
  });
});
