import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { buildScheduleTemplateXlsx } from "@/features/delivery-schedule/lib/buildScheduleTemplateXlsx";
import {
  SECTION_IN_SCHEDULE,
  SECTION_REMAINDER,
} from "@/features/delivery-schedule/lib/scheduleTemplateRows";
import type { BatchDraft, OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";

const plate: OfferPlateForSchedule = {
  id: 1,
  plate_name: "ПБ 60-12-8",
  qty: 40,
};

const batch = (qty: number): BatchDraft => ({
  key: "k1",
  name: "Партия 1",
  deliver_from: "2026-04-01",
  deliver_to: "2026-04-10",
  produce_by: "2026-03-25",
  items: [{ plate_id: 1, qty }],
  status: null,
  ready_date: null,
  hint: null,
  changed: false,
});

const loadSheet = async (blob: Blob) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await blob.arrayBuffer());
  return workbook.getWorksheet("График поставки");
};

describe("buildScheduleTemplateXlsx", () => {
  it("writes a non-empty workbook named График поставки", async () => {
    const blob = await buildScheduleTemplateXlsx([plate], []);

    expect(blob.size).toBeGreaterThan(0);
    const ws = await loadSheet(blob);
    expect(ws).toBeDefined();
    expect(ws?.name).toBe("График поставки");
    expect(ws?.getCell("A1").value).toBe("Партия");
    expect(ws?.getCell("F1").value).toBe("Кол-во");
  });

  it("writes visible section separators and remainder qty", async () => {
    const blob = await buildScheduleTemplateXlsx([plate], [batch(10)]);
    const ws = await loadSheet(blob);

    expect(ws?.getCell("A2").value).toBe(SECTION_IN_SCHEDULE);
    expect(ws?.getCell("A4").value).toBe(SECTION_REMAINDER);
    expect(ws?.getCell("E5").value).toBe("ПБ 60-12-8");
    expect(ws?.getCell("F5").value).toBe(30);
    expect(ws?.getCell("B3").value).toBe("01.04.2026");

    const sectionMerge = ws?.getCell("A2").master === ws?.getCell("F2").master;
    expect(sectionMerge).toBe(true);
    expect(ws?.getCell("A2").fill).toBeTruthy();
  });
});
