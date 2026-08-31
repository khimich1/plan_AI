import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { LayoutBlock } from "@/features/logistics/components/LayoutBlock";
import type { LayoutMetadata, LayoutUnit } from "@/features/logistics/types/logistics";

const unit = (plateName: string, width: number | null = 1.2): LayoutUnit => ({
  completed_plate_id: 1,
  kp_id: 1,
  plate_name: plateName,
  width_m: width,
});

const LAYOUT: LayoutMetadata = {
  body_length_m: 13.2,
  body_used_m: 13.2,
  stacks: [
    {
      index: 1,
      marking_length_m: 8.9,
      tiers: [
        { index: 1, units: [unit("ПБ 89-12-8п"), unit("ПБ 89-12-8п")] },
        { index: 2, units: [unit("ПБ 89-12-8п"), unit("ПБ 80-12-8п")] },
        { index: 3, units: [unit("ПБ 80-12-8п"), unit("ПБ 80-12-8п")] },
      ],
    },
    {
      index: 2,
      marking_length_m: 4.3,
      tiers: [
        { index: 1, units: [unit("ПБ 43-12-8п"), unit("ПБ 42,6-5,3-10п", 0.53)] },
        { index: 2, units: [unit("ПБ 42-3,0-8п", 0.3)] },
      ],
    },
  ],
  loading_steps: [
    { step: 1, stack_index: 1, tier_index: 1, description: "ПБ 89-12-8п ×2" },
    { step: 2, stack_index: 1, tier_index: 2, description: "ПБ 89-12-8п + ПБ 80-12-8п" },
    { step: 3, stack_index: 1, tier_index: 3, description: "ПБ 80-12-8п ×2" },
    { step: 4, stack_index: 2, tier_index: 1, description: "ПБ 43-12-8п + ПБ 42,6-5,3-10п" },
    { step: 5, stack_index: 2, tier_index: 2, description: "ПБ 42-3,0-8п" },
  ],
};

describe("LayoutBlock", () => {
  afterEach(() => {
    cleanup();
  });

  it("ничего не рендерит при layout = null", () => {
    const { container } = render(<LayoutBlock layout={null} />);
    expect(container.firstChild).toBeNull();
  });

  it("ничего не рендерит при пустых штабелях", () => {
    const empty: LayoutMetadata = { ...LAYOUT, stacks: [], loading_steps: [] };
    const { container } = render(<LayoutBlock layout={empty} />);
    expect(container.firstChild).toBeNull();
  });

  it("показывает заголовок с числом штабелей, метражом и весом", () => {
    render(<LayoutBlock layout={LAYOUT} totalWeightKg={19697} maxWeightKg={19800} />);
    const header = screen.getByText(/Укладка в кузов/);
    const text = header.textContent?.replace(/ /g, " ") ?? "";
    expect(text).toContain("2 штабеля");
    expect(text).toContain("13,2 / 13,2 м");
    expect(text).toContain("19 697");
    expect(text).toContain("19 800");
  });

  it("рисует полоску кузова со штабелями и метражом", () => {
    render(<LayoutBlock layout={LAYOUT} />);
    expect(screen.getByText("Штабель 1")).toBeTruthy();
    expect(screen.getByText(/8,9 м/)).toBeTruthy();
    expect(screen.getByText("Штабель 2")).toBeTruthy();
    expect(screen.getByText(/4,3 м/)).toBeTruthy();
  });

  it("по клику раскрывает ярусы штабеля с марками и ширинами", () => {
    render(<LayoutBlock layout={LAYOUT} />);
    expect(screen.queryByText("Ярус 3:")).toBeNull();

    fireEvent.click(screen.getByText("Штабель 1"));
    expect(screen.getByText("Ярус 3:").parentElement?.textContent).toContain("ПБ 80-12-8п (1,2 м)");
    expect(screen.getByText("Ярус 1:").parentElement?.textContent).toContain("ПБ 89-12-8п (1,2 м)");

    fireEvent.click(screen.getByText("Штабель 2"));
    expect(screen.queryByText("Ярус 3:")).toBeNull();
    expect(screen.getByText("Ярус 1:").parentElement?.textContent).toContain("ПБ 42,6-5,3-10п (0,53 м)");
  });

  it("показывает нумерованный порядок погрузки", () => {
    render(<LayoutBlock layout={LAYOUT} />);
    expect(screen.getByText("Порядок погрузки")).toBeTruthy();
    const steps = screen.getAllByRole("listitem");
    expect(steps).toHaveLength(5);
    expect(steps[0].textContent).toBe("Штабель 1, ярус 1: ПБ 89-12-8п ×2");
    expect(steps[4].textContent).toBe("Штабель 2, ярус 2: ПБ 42-3,0-8п");
  });
});
