import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ArchiveSectionTabs } from "@/features/commercial-archive/components/ArchiveSectionTabs";

afterEach(() => {
  cleanup();
});

describe("ArchiveSectionTabs", () => {
  it("shows five sections with approval between archive and production", () => {
    render(<ArchiveSectionTabs value="archived" onChange={vi.fn()} />);

    const tabs = screen.getAllByRole("tab");
    expect(tabs.map((tab) => tab.textContent)).toEqual([
      "📦В архиве",
      "📝На согласовании",
      "🏭В производстве",
      "✅Выполненные",
      "🤝Контрагенты",
    ]);
    expect(tabs[0]).toHaveAttribute("aria-selected", "true");
    expect(tabs[1]).toHaveAttribute("aria-selected", "false");
    expect(tabs[4]).toHaveAttribute("aria-selected", "false");
  });

  it("selects on_approval and does not mark the archive tab", () => {
    const onChange = vi.fn();
    render(<ArchiveSectionTabs value="on_approval" onChange={onChange} />);

    const approval = screen.getByRole("tab", { name: "На согласовании" });
    const archive = screen.getByRole("tab", { name: "В архиве" });
    expect(approval).toHaveAttribute("aria-selected", "true");
    expect(archive).toHaveAttribute("aria-selected", "false");

    fireEvent.click(approval);
    expect(onChange).toHaveBeenCalledWith("on_approval");
  });

  it("selects counterparties and reports that section", () => {
    const onChange = vi.fn();
    render(<ArchiveSectionTabs value="counterparties" onChange={onChange} />);

    const counterparties = screen.getByRole("tab", { name: "Контрагенты" });
    const archive = screen.getByRole("tab", { name: "В архиве" });
    expect(counterparties).toHaveAttribute("aria-selected", "true");
    expect(archive).toHaveAttribute("aria-selected", "false");

    fireEvent.click(counterparties);
    expect(onChange).toHaveBeenCalledWith("counterparties");
  });
});
