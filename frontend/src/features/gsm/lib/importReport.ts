import type { FileImportReport } from "@/features/gsm/types/gsm";

const RECONCILE_EPS = 0.01;

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
