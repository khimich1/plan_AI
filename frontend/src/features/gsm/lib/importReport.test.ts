import { describe, expect, it } from "vitest";
import {
  amountMismatch,
  hasFileReconcileMismatch,
  litersMismatch,
  summarizeImportReport,
} from "@/features/gsm/lib/importReport";
import type { FileImportReport, TransactionImportReport } from "@/features/gsm/types/gsm";

const base = (over: Partial<FileImportReport> = {}): FileImportReport => ({
  filename: "a.xls",
  rows_total: 10,
  rows_inserted: 10,
  rows_duplicate: 0,
  sum_liters: 100,
  sum_amount: 5000,
  footer_liters: 100,
  footer_amount: 5000,
  warnings: [],
  unmatched_cards: [],
  ...over,
});

const report = (
  over: Partial<TransactionImportReport> & { files?: FileImportReport[] },
): TransactionImportReport => ({
  files: over.files ?? [base()],
  rows_inserted: over.rows_inserted ?? 10,
  rows_duplicate: over.rows_duplicate ?? 0,
});

describe("importReport mismatch helpers", () => {
  it("detects liters and amount mismatches against footer", () => {
    const file = base({ sum_liters: 99.5, footer_liters: 100, sum_amount: 5000, footer_amount: 5100 });
    expect(litersMismatch(file)).toBe(true);
    expect(amountMismatch(file)).toBe(true);
    expect(hasFileReconcileMismatch(file)).toBe(true);
  });

  it("treats matching footer as ok", () => {
    const file = base();
    expect(hasFileReconcileMismatch(file)).toBe(false);
  });

  it("flags warning text about totals", () => {
    const file = base({
      warnings: ["a.xls: Кол-во 99.00 ≠ Итоги 100.00"],
    });
    expect(hasFileReconcileMismatch(file)).toBe(true);
  });
});

describe("summarizeImportReport", () => {
  it("success when all rows are new", () => {
    const summary = summarizeImportReport(report({ rows_inserted: 12, rows_duplicate: 0 }));
    expect(summary.tone).toBe("success");
    expect(summary.text).toBe("Добавлено 12 операций.");
    expect(summary.text).not.toMatch(/дубл|вставлен/i);
  });

  it("success with note when some rows already existed", () => {
    const summary = summarizeImportReport(report({ rows_inserted: 12, rows_duplicate: 3 }));
    expect(summary.tone).toBe("success");
    expect(summary.text).toBe(
      "Добавлено 12 операций. Ещё 3 уже были в журнале — повторно не записаны.",
    );
    expect(summary.text).not.toMatch(/дубл|вставлен/i);
  });

  it("info when nothing new — all already in journal", () => {
    const summary = summarizeImportReport(
      report({
        rows_inserted: 0,
        rows_duplicate: 3,
        files: [base({ rows_total: 3, rows_inserted: 0, rows_duplicate: 3 })],
      }),
    );
    expect(summary.tone).toBe("info");
    expect(summary.text).toBe(
      "Новых операций нет: все 3 уже есть в журнале. Повторная загрузка того же файла ничего не меняет.",
    );
    expect(summary.text).not.toMatch(/дубл|вставлен/i);
  });

  it("warning when footer totals diverge", () => {
    const summary = summarizeImportReport(
      report({
        rows_inserted: 12,
        rows_duplicate: 1,
        files: [
          base({ rows_inserted: 10, rows_duplicate: 0 }),
          base({
            filename: "bad.xls",
            rows_total: 5,
            rows_inserted: 2,
            rows_duplicate: 1,
            sum_liters: 40,
            footer_liters: 45,
            warnings: ["bad.xls: Кол-во 40.00 ≠ Итоги 45.00"],
          }),
        ],
      }),
    );
    expect(summary.tone).toBe("warning");
    expect(summary.text).toContain("Добавлено 12 операций");
    expect(summary.text).toContain("Ещё 1 уже были в журнале");
    expect(summary.text).toContain("Расхождение итогов по 1 файлам.");
    expect(summary.text).not.toMatch(/дубл|вставлен/i);
  });
});
