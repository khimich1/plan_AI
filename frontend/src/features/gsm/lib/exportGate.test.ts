import { describe, expect, it } from "vitest";
import {
  exportConfirmMessages,
  exportDisabledReason,
  exportHardBlock,
  hasExportedDays,
  hasSoftExportWarnings,
} from "@/features/gsm/lib/exportGate";
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

describe("exportHardBlock", () => {
  it("blocks when a day has manual_intervention and lists the date", () => {
    const result = exportHardBlock([
      base({ date: "2025-04-07", warnings: ["manual_intervention"] }),
      base({ id: 2, date: "2025-04-08", warnings: ["balance_route"] }),
    ]);
    expect(result.blocked).toBe(true);
    expect(result.dates).toEqual(["2025-04-07"]);
  });

  it("blocks on day-level unsolvable", () => {
    const result = exportHardBlock([base({ warnings: ["unsolvable"] })]);
    expect(result.blocked).toBe(true);
  });

  it("blocks on period warning unsolvable even without day dates", () => {
    const result = exportHardBlock([base({})], ["unsolvable"]);
    expect(result.blocked).toBe(true);
    expect(result.dates).toEqual([]);
  });

  it("does not block clean draft days", () => {
    expect(exportHardBlock([base({})]).blocked).toBe(false);
  });
});

describe("hasSoftExportWarnings", () => {
  it("is true for yellow day warnings when period is exportable", () => {
    expect(hasSoftExportWarnings([base({ warnings: ["weekend_anchor"] })])).toBe(true);
    expect(hasSoftExportWarnings([base({ warnings: ["hook_above_threshold"] })])).toBe(true);
    expect(hasSoftExportWarnings([base({ warnings: ["balance_route"] })])).toBe(true);
  });

  it("is false when hard-blocked even if yellow warnings exist", () => {
    expect(
      hasSoftExportWarnings([
        base({ date: "2025-04-07", warnings: ["manual_intervention", "balance_route"] }),
      ]),
    ).toBe(false);
  });

  it("is false for clean days", () => {
    expect(hasSoftExportWarnings([base({})])).toBe(false);
  });
});

describe("exportConfirmMessages", () => {
  it("is empty for clean drafts", () => {
    expect(exportConfirmMessages([base({})])).toEqual([]);
  });

  it("asks about yellow warnings and re-export independently", () => {
    expect(exportConfirmMessages([base({ warnings: ["weekend_anchor"] })])).toEqual([
      "В периоде есть предупреждения. Экспортировать всё равно?",
    ]);
    expect(exportConfirmMessages([base({ status: "exported" })])).toEqual([
      "Период уже экспортировался. Скачать снова?",
    ]);
    expect(
      exportConfirmMessages([base({ status: "exported", warnings: ["balance_route"] })]),
    ).toEqual([
      "В периоде есть предупреждения. Экспортировать всё равно?",
      "Период уже экспортировался. Скачать снова?",
    ]);
  });
});

describe("hasExportedDays", () => {
  it("detects exported status in the period", () => {
    expect(hasExportedDays([base({ status: "exported" })])).toBe(true);
    expect(hasExportedDays([base({ status: "draft" })])).toBe(false);
    expect(hasExportedDays([base({ status: "confirmed" })])).toBe(false);
  });
});

describe("exportDisabledReason", () => {
  it("explains empty period", () => {
    expect(exportDisabledReason([])).toBe("Нет путевых листов за период.");
  });

  it("lists manual days to fix", () => {
    expect(
      exportDisabledReason([
        base({ date: "2025-04-07", warnings: ["manual_intervention"] }),
        base({ id: 2, date: "2025-04-08", warnings: ["manual_intervention"] }),
      ]),
    ).toBe("Исправьте дни ручной доработки: 2025-04-07, 2025-04-08.");
  });

  it("explains unsolvable period without dates", () => {
    expect(exportDisabledReason([base({})], ["unsolvable"])).toBe(
      "Период нерешаем — исправьте дни вручную.",
    );
  });

  it("is null when export is allowed", () => {
    expect(exportDisabledReason([base({ warnings: ["weekend_anchor"] })])).toBeNull();
  });
});
