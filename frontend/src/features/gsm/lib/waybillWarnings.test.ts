import { describe, expect, it } from "vitest";
import {
  formatRouteSummary,
  isAnchorDay,
  isProblematicDay,
  warningDetailText,
  warningMeta,
} from "@/features/gsm/lib/waybillWarnings";
import type { GsmWaybill } from "@/features/gsm/types/gsm";

const base = (overrides: Partial<GsmWaybill>): GsmWaybill => ({
  id: 1,
  vehicle_id: 1,
  date: "2025-04-01",
  driver_id: 1,
  status: "draft",
  source: "auto",
  odometer_start: 0,
  odometer_end: 100,
  fuel_start: 10,
  fuel_issued: 0,
  fuel_end: 5,
  km: 100,
  route: [],
  warnings: [],
  ...overrides,
});

describe("waybillWarnings helpers", () => {
  it("maps known warning codes to Russian reason text", () => {
    expect(warningMeta("weekend_anchor").short).toBe("Выходной");
    expect(warningMeta("hook_above_threshold").reason).toMatch(/порог/i);
    expect(warningMeta("unsolvable").short).toBe("Нерешаемо");
  });

  it("maps manual_intervention and balance_route to short labels and reasons", () => {
    expect(warningMeta("manual_intervention").short).toBe("Ручная доработка");
    expect(warningMeta("manual_intervention").reason).toMatch(/баланс|ручн/i);
    expect(warningMeta("balance_route").short).toBe("Маршрут для баланса");
    expect(warningMeta("balance_route").reason).toMatch(/удлин|маршрут/i);
  });

  it("maps borrowed_route to a short yellow label about another vehicle", () => {
    expect(warningMeta("borrowed_route").short).toBe("Чужой пул");
    expect(warningMeta("borrowed_route").reason).toMatch(/друг|машин/i);
  });

  it("detects anchor days by fuel issued, warnings or station", () => {
    expect(isAnchorDay(base({ fuel_issued: 30 }))).toBe(true);
    expect(isAnchorDay(base({ warnings: ["weekend_anchor"] }))).toBe(true);
    expect(isAnchorDay(base({ route: [{ from: "A", to: "B", km: 10, station_id: 3 }] }))).toBe(true);
    expect(isAnchorDay(base({}))).toBe(false);
  });

  it("prefers warning_details.detail over generic reason", () => {
    const wb = base({
      warnings: ["manual_intervention"],
      warning_details: [{ code: "manual_intervention", detail: "бак не сходится" }],
    });
    expect(warningDetailText(wb, "manual_intervention")).toBe("бак не сходится");
    expect(warningDetailText(base({ warnings: ["manual_intervention"] }), "manual_intervention")).toMatch(
      /ручн/i,
    );
  });

  it("detects problematic days by manual_intervention warning", () => {
    expect(isProblematicDay(base({ warnings: ["manual_intervention"] }))).toBe(true);
    expect(isProblematicDay(base({ warnings: ["balance_route"] }))).toBe(false);
    expect(isProblematicDay(base({ warnings: ["borrowed_route"] }))).toBe(false);
    expect(isProblematicDay(base({}))).toBe(false);
  });

  it("formats route summary", () => {
    expect(formatRouteSummary([{ from: "A", to: "B", km: 10 }])).toBe("A → B");
    expect(
      formatRouteSummary([
        { from: "A", to: "B", km: 10 },
        { from: "B", to: "C", km: 20 },
      ]),
    ).toMatch(/2 плеч/);
  });
});
