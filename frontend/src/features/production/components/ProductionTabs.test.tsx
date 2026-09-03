import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ProductionTabs } from "@/features/production/components/ProductionTabs";
import type { ProductionTab } from "@/features/production/types/production";

describe("ProductionTabs", () => {
  afterEach(() => {
    cleanup();
  });

  it("places КП в работе immediately after Планы", () => {
    const onChange = vi.fn();
    render(<ProductionTabs value={"calendar" as ProductionTab} onChange={onChange} />);
    const tabs = screen.getAllByRole("tab").map((tab) => tab.textContent ?? "");
    const plans = tabs.findIndex((label) => label.includes("Планы"));
    const inWork = tabs.findIndex((label) => label.includes("КП в работе"));
    expect(plans).toBeGreaterThanOrEqual(0);
    expect(inWork).toBe(plans + 1);
  });
});
