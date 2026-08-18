import { describe, expect, it } from "vitest";
import {
  amountMismatch,
  hasFileReconcileMismatch,
  litersMismatch,
} from "@/features/gsm/lib/importReport";
import type { FileImportReport } from "@/features/gsm/types/gsm";

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
