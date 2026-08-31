import ExcelJS from "exceljs";
import type { BatchDraft, OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";
import {
  buildScheduleTemplateRows,
  type TemplateRow,
} from "@/features/delivery-schedule/lib/scheduleTemplateRows";

export const TEMPLATE_FILENAME = "delivery_schedule_template.xlsx";
export const TEMPLATE_SHEET_NAME = "График поставки";

const HEADERS = ["Партия", "Поставка с", "Поставка по", "Произвести до", "Марка", "Кол-во"] as const;
const COLUMN_WIDTHS = [28, 14, 14, 14, 28, 10] as const;
const SECTION_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFD9E2F3" },
};
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const itemValues = (row: Extract<TemplateRow, { kind: "item" }>): (string | number)[] => [
  row.batchName,
  row.deliverFrom,
  row.deliverTo,
  row.produceBy,
  row.plateName,
  row.qty,
];

export async function buildScheduleTemplateXlsx(
  plates: OfferPlateForSchedule[],
  batches: BatchDraft[],
): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet(TEMPLATE_SHEET_NAME);

  ws.columns = COLUMN_WIDTHS.map((width) => ({ width }));

  const header = ws.addRow([...HEADERS]);
  header.font = { bold: true };
  header.alignment = { wrapText: true, vertical: "top" };

  for (const row of buildScheduleTemplateRows(plates, batches)) {
    if (row.kind === "section") {
      const excelRow = ws.addRow([row.title]);
      ws.mergeCells(excelRow.number, 1, excelRow.number, HEADERS.length);
      const cell = excelRow.getCell(1);
      cell.fill = SECTION_FILL;
      cell.font = { bold: true };
      cell.alignment = { vertical: "middle" };
      continue;
    }
    ws.addRow(itemValues(row));
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: XLSX_MIME });
}
