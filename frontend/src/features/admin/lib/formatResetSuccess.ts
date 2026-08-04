import type { DbResetReport, ResetVariant } from "@/features/admin/types/admin";

export function formatResetSuccess(
  report: DbResetReport,
  variant: ResetVariant,
): string {
  const parts: string[] = [];
  const sqlite = report.sqlite ?? {};
  const plans = report.plans ?? {};

  if (variant === "full" || variant === "kp-only") {
    const kpOffers = sqlite.kp_offers ?? 0;
    if (kpOffers > 0) {
      parts.push(`${kpOffers} КП`);
    }
  }

  if (variant === "full" || variant === "plans-only") {
    const sqlitePlans = plans.sqlite_plans ?? 0;
    const legacyPlans =
      (plans.plan_files ?? 0) +
      (plans.current_plan ?? 0) +
      (plans.metadata ?? 0);
    const planTotal = sqlitePlans + legacyPlans;
    if (planTotal > 0) {
      parts.push(`${planTotal} планов`);
    }
  }

  if (variant === "full") {
    const archivedLegacy =
      (plans.archived_plan_files ?? 0) +
      (plans.archived_metadata ?? 0) +
      (plans.archived_calendar ?? 0);
    if (archivedLegacy > 0) {
      parts.push(`${archivedLegacy} legacy JSON в архиве`);
    }
  }

  if ((variant === "full" || variant === "calendar-only") && report.calendar_reset) {
    parts.push("календарь сброшен");
  }

  if (parts.length === 0) {
    return "Операция выполнена успешно.";
  }

  return `Удалено: ${parts.join(", ")}.`;
}
