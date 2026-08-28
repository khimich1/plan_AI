import type { FileImportReport, TransactionImportReport } from "@/features/gsm/types/gsm";

const RECONCILE_EPS = 0.01;

export type ImportSummaryTone = "info" | "warning" | "success";

export type ImportSummary = {
  tone: ImportSummaryTone;
  text: string;
};

export const litersMismatch = (file: FileImportReport): boolean =>
  file.footer_liters != null && Math.abs(file.sum_liters - file.footer_liters) > RECONCILE_EPS;

export const amountMismatch = (file: FileImportReport): boolean =>
  file.footer_amount != null && Math.abs(file.sum_amount - file.footer_amount) > RECONCILE_EPS;

/** True when footer totals diverge from parsed sums or parser left reconcile warnings. */
export const hasFileReconcileMismatch = (file: FileImportReport): boolean => {
  if (litersMismatch(file) || amountMismatch(file)) {
    return true;
  }
  return file.warnings.some(
    (w) => w.includes("≠") || w.toLowerCase().includes("итог"),
  );
};

export const formatLiters = (value: number | null): string =>
  value == null ? "—" : value.toLocaleString("ru-RU", { maximumFractionDigits: 2 });

export const formatAmount = (value: number | null): string =>
  value == null ? "—" : value.toLocaleString("ru-RU", { maximumFractionDigits: 2 });

/** Human-readable import outcome for accountant UI (no «дубль» / «вставлено»). */
export const summarizeImportReport = (report: TransactionImportReport): ImportSummary => {
  const mismatchCount = report.files.filter(hasFileReconcileMismatch).length;
  const inserted = report.rows_inserted;
  const already = report.rows_duplicate;

  let text: string;
  let tone: ImportSummaryTone;

  if (inserted === 0 && already > 0) {
    tone = "info";
    text =
      `Новых операций нет: все ${already} уже есть в журнале. ` +
      "Повторная загрузка того же файла ничего не меняет.";
  } else if (inserted > 0 && already > 0) {
    tone = "success";
    text =
      `Добавлено ${inserted} операций. ` +
      `Ещё ${already} уже были в журнале — повторно не записаны.`;
  } else if (inserted > 0) {
    tone = "success";
    text = `Добавлено ${inserted} операций.`;
  } else {
    tone = "info";
    text = "Новых операций нет.";
  }

  if (mismatchCount > 0) {
    tone = "warning";
    text = `${text} Расхождение итогов по ${mismatchCount} файлам.`;
  }

  return { tone, text };
};
