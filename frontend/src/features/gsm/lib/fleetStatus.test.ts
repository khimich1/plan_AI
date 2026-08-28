import { describe, expect, it } from "vitest";
import {
  currentMonthBounds,
  fleetOpenBeforeMonth,
  fleetStatusMeta,
  litersDiffHidden,
  litersDiffOk,
  openBeforeMonthLabel,
  openBeforeSummary,
} from "@/features/gsm/lib/fleetStatus";

describe("fleetStatusMeta", () => {
  it("maps all six statuses to Russian labels and tones", () => {
    expect(fleetStatusMeta("no_data")).toEqual({ label: "Нет данных", tone: "muted" });
    expect(fleetStatusMeta("needs_generation")).toEqual({
      label: "Требуется генерация",
      tone: "warning",
    });
    expect(fleetStatusMeta("has_red_days")).toEqual({ label: "Есть красные дни", tone: "danger" });
    expect(fleetStatusMeta("drafts_pending")).toEqual({ label: "Черновики", tone: "warning" });
    expect(fleetStatusMeta("pending_export")).toEqual({ label: "Готово к экспорту", tone: "info" });
    expect(fleetStatusMeta("ready")).toEqual({ label: "Выгружено", tone: "success" });
  });
});

describe("litersDiff helpers", () => {
  it("hides the badge when there are no waybills", () => {
    expect(litersDiffHidden(0)).toBe(true);
    expect(litersDiffHidden(1)).toBe(false);
  });

  it("treats |Δ| ≤ 0.01 as balanced", () => {
    expect(litersDiffOk(0)).toBe(true);
    expect(litersDiffOk(0.01)).toBe(true);
    expect(litersDiffOk(-0.01)).toBe(true);
    expect(litersDiffOk(0.02)).toBe(false);
    expect(litersDiffOk(-2)).toBe(false);
  });
});

describe("openBeforeSummary", () => {
  it("is empty when every row has open_before = 0", () => {
    expect(openBeforeSummary([{ open_before: 0 }, { open_before: 0 }])).toEqual({
      pl: 0,
      vehicles: 0,
    });
  });

  it("sums waybills and counts vehicles with a tail", () => {
    expect(openBeforeSummary([{ open_before: 2 }, { open_before: 0 }, { open_before: 1 }])).toEqual({
      pl: 3,
      vehicles: 2,
    });
  });
});

describe("currentMonthBounds", () => {
  it("returns the calendar month of the given date", () => {
    expect(currentMonthBounds(new Date(2026, 7, 24))).toEqual({
      from: "2026-08-01",
      to: "2026-08-31",
    });
  });
});

describe("fleetOpenBeforeMonth", () => {
  it("returns null when no row has a tail", () => {
    expect(
      fleetOpenBeforeMonth([
        { open_before: 0, open_before_month: null },
        { open_before: 0, open_before_month: "2026-07" },
      ]),
    ).toBeNull();
  });

  it("picks the max open_before_month among rows with a tail", () => {
    expect(
      fleetOpenBeforeMonth([
        { open_before: 2, open_before_month: "2026-06" },
        { open_before: 0, open_before_month: null },
        { open_before: 6, open_before_month: "2026-07" },
      ]),
    ).toBe("2026-07");
  });
});

describe("openBeforeMonthLabel", () => {
  it("names July in Russian for 2026-07", () => {
    const label = openBeforeMonthLabel("2026-07");
    expect(label.toLowerCase()).toContain("июл");
  });

  it("capitalizes the month for «Июль не выгружен»", () => {
    expect(openBeforeMonthLabel("2026-07").startsWith("И")).toBe(true);
  });
});
