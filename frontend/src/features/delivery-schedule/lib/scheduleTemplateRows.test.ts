import { describe, expect, it } from "vitest";
import {
  SECTION_IN_SCHEDULE,
  SECTION_REMAINDER,
  buildScheduleTemplateRows,
  type TemplateRow,
} from "@/features/delivery-schedule/lib/scheduleTemplateRows";
import type { BatchDraft, OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";

const plate = (
  overrides: Partial<OfferPlateForSchedule> & Pick<OfferPlateForSchedule, "id" | "plate_name" | "qty">,
): OfferPlateForSchedule => ({
  position_number: overrides.position_number ?? overrides.id,
  ...overrides,
});

const batch = (overrides: Partial<BatchDraft> & { items: BatchDraft["items"] }): BatchDraft => ({
  key: "k1",
  name: "Партия 1",
  deliver_from: "2026-04-01",
  deliver_to: "2026-04-10",
  produce_by: "2026-03-25",
  status: null,
  ready_date: null,
  hint: null,
  changed: false,
  ...overrides,
});

const itemRows = (rows: TemplateRow[]) => rows.filter((row) => row.kind === "item");

describe("buildScheduleTemplateRows", () => {
  it("empty draft: only remainder section with KP qty", () => {
    const plates = [plate({ id: 1, plate_name: "ПБ 60-12-8", qty: 40 })];

    const rows = buildScheduleTemplateRows(plates, []);

    expect(rows).toEqual([
      { kind: "section", title: SECTION_REMAINDER },
      {
        kind: "item",
        batchName: "",
        deliverFrom: "",
        deliverTo: "",
        produceBy: "",
        plateName: "ПБ 60-12-8",
        qty: 40,
      },
    ]);
  });

  it("partial draft: scheduled band then remainder", () => {
    const plates = [plate({ id: 1, plate_name: "ПБ 60-12-8", qty: 40 })];
    const batches = [batch({ items: [{ plate_id: 1, qty: 10 }] })];

    const rows = buildScheduleTemplateRows(plates, batches);

    expect(rows).toEqual([
      { kind: "section", title: SECTION_IN_SCHEDULE },
      {
        kind: "item",
        batchName: "Партия 1",
        deliverFrom: "01.04.2026",
        deliverTo: "10.04.2026",
        produceBy: "25.03.2026",
        plateName: "ПБ 60-12-8",
        qty: 10,
      },
      { kind: "section", title: SECTION_REMAINDER },
      {
        kind: "item",
        batchName: "",
        deliverFrom: "",
        deliverTo: "",
        produceBy: "",
        plateName: "ПБ 60-12-8",
        qty: 30,
      },
    ]);
  });

  it("two batches of the same mark stay as two top rows", () => {
    const plates = [plate({ id: 1, plate_name: "ПБ 60-12-8", qty: 40 })];
    const batches = [
      batch({ items: [{ plate_id: 1, qty: 10 }] }),
      batch({ key: "k2", name: "Партия 2", items: [{ plate_id: 1, qty: 15 }] }),
    ];

    const rows = buildScheduleTemplateRows(plates, batches);
    const top = itemRows(rows).filter((_, index) => index < 2);

    expect(rows[0]).toEqual({ kind: "section", title: SECTION_IN_SCHEDULE });
    expect(top.map((row) => row.qty)).toEqual([10, 15]);
    expect(rows.at(-1)).toMatchObject({ kind: "item", qty: 15, plateName: "ПБ 60-12-8" });
  });

  it("fully allocated: remainder section is omitted", () => {
    const plates = [plate({ id: 1, plate_name: "ПБ 60-12-8", qty: 40 })];
    const batches = [batch({ items: [{ plate_id: 1, qty: 40 }] })];

    const rows = buildScheduleTemplateRows(plates, batches);

    expect(rows.some((row) => row.kind === "section" && row.title === SECTION_REMAINDER)).toBe(
      false,
    );
    expect(itemRows(rows)).toHaveLength(1);
    expect(itemRows(rows)[0].qty).toBe(40);
  });

  it("skips plates without a mark and items with qty < 1", () => {
    const plates = [
      plate({ id: 1, plate_name: "   ", qty: 8 }),
      plate({ id: 2, plate_name: "ПБ 72-15-8", qty: 10 }),
    ];
    const batches = [
      batch({
        items: [
          { plate_id: 1, qty: 3 },
          { plate_id: 2, qty: 0 },
          { plate_id: 2, qty: 4 },
        ],
      }),
    ];

    const rows = buildScheduleTemplateRows(plates, batches);

    expect(itemRows(rows).map((row) => ({ plateName: row.plateName, qty: row.qty }))).toEqual([
      { plateName: "ПБ 72-15-8", qty: 4 },
      { plateName: "ПБ 72-15-8", qty: 6 },
    ]);
  });
});
