import { describe, expect, it } from "vitest";
import {
  bulkExportConfirmMessages,
  bulkKitConfirmMessages,
  exportConfirmMessages,
  exportDisabledReason,
  exportHardBlock,
  hasExportedDays,
  hasSoftExportWarnings,
  planBulkExport,
  planBulkGenerate,
  planKit,
} from "@/features/gsm/lib/exportGate";
import type { FleetOverviewRow, GsmWaybill } from "@/features/gsm/types/gsm";

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

  it("does not hard-stop a day with only borrowed_route", () => {
    expect(exportHardBlock([base({ warnings: ["borrowed_route"] })]).blocked).toBe(false);
  });
});

describe("hasSoftExportWarnings", () => {
  it("is true for yellow day warnings when period is exportable", () => {
    expect(hasSoftExportWarnings([base({ warnings: ["weekend_anchor"] })])).toBe(true);
    expect(hasSoftExportWarnings([base({ warnings: ["hook_above_threshold"] })])).toBe(true);
    expect(hasSoftExportWarnings([base({ warnings: ["balance_route"] })])).toBe(true);
    expect(hasSoftExportWarnings([base({ warnings: ["borrowed_route"] })])).toBe(true);
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

  it("does not disable zip for borrowed_route alone; still blocks manual_intervention", () => {
    expect(exportDisabledReason([base({ warnings: ["borrowed_route"] })])).toBeNull();
    expect(
      exportDisabledReason([base({ date: "2025-04-07", warnings: ["manual_intervention"] })]),
    ).toBe("Исправьте дни ручной доработки: 2025-04-07.");
  });
});

const overview = (
  id: number,
  overrides: Partial<FleetOverviewRow> = {},
): FleetOverviewRow => ({
  vehicle: { id, name: `Машина ${id}`, plate_number: `A${id}` },
  tx_count: 1,
  tx_liters: 10,
  tx_amount: 100,
  tx_last_date: "2026-08-10",
  wb_count: 1,
  wb_km: 100,
  wb_fuel_issued: 10,
  wb_last_date: "2026-08-10",
  red_days: 0,
  draft_count: 0,
  confirmed_count: 0,
  exported_count: 0,
  fuel_end_last: 20,
  liters_diff: 0,
  open_before: 0,
  open_before_month: null,
  chain_broken: false,
  status: "ready",
  ...overrides,
});

describe("planBulkExport", () => {
  it("excludes red_days vehicles and keeps id order for the zip", () => {
    const plan = planBulkExport(
      [
        overview(3, { red_days: 0 }),
        overview(1, { red_days: 2, status: "has_red_days" }),
        overview(2, { red_days: 0, exported_count: 1 }),
      ],
      [2, 1, 3],
    );
    expect(plan.excluded).toEqual([
      {
        vehicleId: 1,
        label: "Машина 1 (A1)",
        reason: "Исправьте дни ручной доработки (2).",
      },
    ]);
    expect(plan.cleanIds).toEqual([2, 3]);
    expect(plan.alreadyExportedCount).toBe(1);
    expect(plan.cleanCount).toBe(2);
  });
});

describe("bulkExportConfirmMessages", () => {
  it("summarizes already-exported vehicles without per-day details", () => {
    expect(
      bulkExportConfirmMessages({
        cleanIds: [2, 3],
        excluded: [],
        alreadyExportedCount: 1,
        cleanCount: 2,
      }),
    ).toEqual(["1 из 2 уже экспортировались. Скачать снова?"]);
    expect(
      bulkExportConfirmMessages({
        cleanIds: [1],
        excluded: [],
        alreadyExportedCount: 0,
        cleanCount: 1,
      }),
    ).toEqual([]);
  });
});

describe("planKit", () => {
  it("excludes a July tail when the selected period is August", () => {
    const plan = planKit(
      [
        overview(1, { open_before: 6, open_before_month: "2026-07" }),
        overview(2),
      ],
      null,
      "2026-08-01",
    );
    expect(plan.cleanIds).not.toContain(1);
    expect(plan.cleanIds).toContain(2);
    expect(plan.excluded).toEqual([
      {
        vehicleId: 1,
        label: "Машина 1 (A1)",
        reason: "Сначала выгрузите Июль",
      },
    ]);
  });

  it("includes the same tail row when periodFrom is the tail month and it is not red", () => {
    const plan = planKit(
      [overview(1, { open_before: 6, open_before_month: "2026-07" })],
      null,
      "2026-07-01",
    );
    expect(plan.cleanIds).toEqual([1]);
    expect(plan.excluded).toEqual([]);
  });

  it("still excludes a red row when periodFrom is the tail month", () => {
    const plan = planKit(
      [overview(1, { open_before: 6, open_before_month: "2026-07", red_days: 2 })],
      null,
      "2026-07-01",
    );
    expect(plan.cleanIds).toEqual([]);
    expect(plan.excluded[0]?.reason).toMatch(/ручной доработки/);
  });

  it("prefers the red reason when a row is both red and has a tail in another month", () => {
    const plan = planKit(
      [
        overview(1, {
          open_before: 6,
          open_before_month: "2026-07",
          red_days: 2,
          status: "has_red_days",
        }),
      ],
      null,
      "2026-08-01",
    );
    expect(plan.cleanIds).toEqual([]);
    expect(plan.excluded).toEqual([
      {
        vehicleId: 1,
        label: "Машина 1 (A1)",
        reason: "Исправьте дни ручной доработки (2).",
      },
    ]);
  });

  it("excludes chain_broken on the current period with a non-red reason", () => {
    const plan = planKit(
      [overview(1, { chain_broken: true, red_days: 0 })],
      null,
      "2026-08-01",
    );
    expect(plan.cleanIds).not.toContain(1);
    expect(plan.excluded).toHaveLength(1);
    expect(plan.excluded[0]?.reason).not.toMatch(/ручной доработки/);
    expect(plan.excluded[0]?.reason).toMatch(/^Пересчитайте Август/);
  });

  it("excludes chain_broken even when the period is the tail month", () => {
    const plan = planKit(
      [
        overview(1, {
          open_before: 6,
          open_before_month: "2026-07",
          chain_broken: true,
        }),
      ],
      null,
      "2026-07-01",
    );
    expect(plan.cleanIds).toEqual([]);
    expect(plan.excluded[0]?.reason).toMatch(/^Пересчитайте Июль/);
  });

  it("looks at all rows when selectedIds is null", () => {
    const plan = planKit([overview(1), overview(2), overview(3)], null, "2026-08-01");
    expect(plan.cleanIds).toEqual([1, 2, 3]);
  });

  it("looks at no rows when selectedIds is empty", () => {
    const plan = planKit([overview(1), overview(2)], [], "2026-08-01");
    expect(plan.cleanIds).toEqual([]);
    expect(plan.excluded).toEqual([]);
  });

  it("looks only at selected ids when a list is given", () => {
    const plan = planKit([overview(1), overview(2), overview(3)], [3, 1], "2026-08-01");
    expect(plan.cleanIds).toEqual([1, 3]);
  });
});

describe("bulkKitConfirmMessages", () => {
  it("asks to re-download when clean kit rows are already fully exported", () => {
    const plan = planKit(
      [
        overview(1, { wb_count: 4, exported_count: 4 }),
        overview(2, { wb_count: 3, exported_count: 1 }),
      ],
      null,
      "2026-08-01",
    );
    expect(plan.alreadyExportedCount).toBe(1);
    expect(plan.cleanCount).toBe(2);
    expect(bulkKitConfirmMessages(plan)).toEqual([
      "1 из 2 уже экспортировались. Скачать снова?",
    ]);
  });

  it("does not count fully exported excluded rows toward alreadyExportedCount", () => {
    const plan = planKit(
      [
        overview(1, { wb_count: 4, exported_count: 4, red_days: 1 }),
        overview(2, { wb_count: 3, exported_count: 0 }),
      ],
      null,
      "2026-08-01",
    );
    expect(plan.alreadyExportedCount).toBe(0);
    expect(plan.cleanCount).toBe(1);
    expect(bulkKitConfirmMessages(plan)).toEqual([]);
  });

  it("asks to re-download when every clean kit row is already fully exported", () => {
    const plan = planKit(
      [
        overview(1, { wb_count: 2, exported_count: 2 }),
        overview(2, { wb_count: 5, exported_count: 5 }),
      ],
      null,
      "2026-08-01",
    );
    expect(plan.alreadyExportedCount).toBe(2);
    expect(plan.cleanCount).toBe(2);
    expect(bulkKitConfirmMessages(plan)).toEqual([
      "2 из 2 уже экспортировались. Скачать снова?",
    ]);
  });
});

describe("planBulkGenerate", () => {
  it("skips a July tail on August and keeps a clean neighbor", () => {
    const plan = planBulkGenerate(
      [
        overview(1, { open_before: 6, open_before_month: "2026-07" }),
        overview(2),
      ],
      [1, 2],
      "2026-08-01",
    );
    expect(plan.eligibleIds).toEqual([2]);
    expect(plan.skipped).toEqual([
      {
        vehicleId: 1,
        label: "Машина 1 (A1)",
        reason: "сначала выгрузите Июль",
      },
    ]);
  });

  it("allows generate when the period is the tail month", () => {
    const plan = planBulkGenerate(
      [overview(1, { open_before: 6, open_before_month: "2026-07", chain_broken: true })],
      [1],
      "2026-07-01",
    );
    expect(plan.eligibleIds).toEqual([1]);
    expect(plan.skipped).toEqual([]);
  });

  it("skips chain_broken on the current period", () => {
    const plan = planBulkGenerate(
      [overview(1, { chain_broken: true })],
      [1],
      "2026-08-01",
    );
    expect(plan.eligibleIds).toEqual([]);
    expect(plan.skipped[0]?.reason).toBe("пересчитайте Август");
  });
});
