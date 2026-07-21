import type { QueryClient } from "@tanstack/react-query";

import { productionApi } from "@/features/production/api/productionApi";
import {
  ApiError,
  getPlanVersionConflictDetails,
  isPlanVersionConflict as isPlanVersionConflictFromApiError,
  PLAN_VERSION_CONFLICT_CODE,
} from "@/shared/lib/apiError";

export { PLAN_VERSION_CONFLICT_CODE };

export const PLAN_VERSION_CONFLICT_TOAST_MESSAGE =
  "План был изменён. Данные обновлены — повторите действие.";

export const isPlanVersionConflict = isPlanVersionConflictFromApiError;

const productionQueryKeys = {
  all: ["production"] as const,
  plan: (planId: string) => ["production", "plan", planId] as const,
  day: (date: string) => ["production", "day", date] as const,
};

const TOAST_CONTAINER_ID = "plan-conflict-toast-root";
const TOAST_DEBOUNCE_MS = 1_000;
const TOAST_VISIBLE_MS = 5_000;

let lastToastAt = 0;
let reloadInFlight: Promise<boolean> | null = null;

const resolvePlanIdFromVariables = (variables: unknown): string | undefined => {
  if (!variables || typeof variables !== "object") {
    return undefined;
  }
  const record = variables as Record<string, unknown>;
  if (typeof record.planId === "string") {
    return record.planId;
  }
  if (typeof record.active_plan_id === "string") {
    return record.active_plan_id;
  }
  return undefined;
};

const resolveDateFromVariables = (variables: unknown): string | undefined => {
  if (!variables || typeof variables !== "object") {
    return undefined;
  }
  const record = variables as Record<string, unknown>;
  return typeof record.date === "string" ? record.date : undefined;
};

export const showPlanConflictToast = (
  message: string = PLAN_VERSION_CONFLICT_TOAST_MESSAGE,
): void => {
  if (typeof document === "undefined") {
    return;
  }

  const now = Date.now();
  if (now - lastToastAt < TOAST_DEBOUNCE_MS) {
    return;
  }
  lastToastAt = now;

  let container = document.getElementById(TOAST_CONTAINER_ID);
  if (!container) {
    container = document.createElement("div");
    container.id = TOAST_CONTAINER_ID;
    container.style.cssText =
      "position:fixed;bottom:1.5rem;right:1.5rem;z-index:10000;display:flex;flex-direction:column;gap:0.5rem;pointer-events:none;max-width:min(22rem,calc(100vw - 2rem));";
    document.body.appendChild(container);
  }

  const toast = document.createElement("div");
  toast.textContent = message;
  toast.setAttribute("role", "status");
  toast.style.cssText =
    "background:#101828;color:#ffffff;padding:0.75rem 1rem;border-radius:12px;box-shadow:0 8px 24px rgba(16,24,40,0.18);font-size:0.92rem;line-height:1.4;";
  container.appendChild(toast);

  window.setTimeout(() => {
    toast.remove();
    if (container && container.childElementCount === 0) {
      container.remove();
    }
  }, TOAST_VISIBLE_MS);
};

export type HandlePlanVersionConflictOptions = {
  planId?: string;
  date?: string;
  variables?: unknown;
  showToast?: boolean;
};

export const handlePlanVersionConflict = async (
  queryClient: QueryClient,
  error: unknown,
  options: HandlePlanVersionConflictOptions = {},
): Promise<boolean> => {
  if (!isPlanVersionConflict(error)) {
    return false;
  }

  if (reloadInFlight) {
    return reloadInFlight;
  }

  reloadInFlight = (async () => {
    if (options.showToast !== false) {
      showPlanConflictToast();
    }

    const details = getPlanVersionConflictDetails(error);
    const planId =
      options.planId ??
      details?.plan_id ??
      resolvePlanIdFromVariables(options.variables);
    const date =
      options.date ?? resolveDateFromVariables(options.variables);

    await queryClient.invalidateQueries({
      queryKey: productionQueryKeys.all,
      refetchType: "active",
    });

    if (planId) {
      await queryClient.fetchQuery({
        queryKey: productionQueryKeys.plan(planId),
        queryFn: () => productionApi.getPlan(planId),
      });
    }

    if (date) {
      await queryClient.invalidateQueries({
        queryKey: productionQueryKeys.day(date),
        refetchType: "active",
      });
    }

    return true;
  })();

  try {
    return await reloadInFlight;
  } finally {
    reloadInFlight = null;
  }
};

export const __testables = {
  resetPlanConflictUiState: () => {
    lastToastAt = 0;
    reloadInFlight = null;
    document.getElementById(TOAST_CONTAINER_ID)?.remove();
  },
  isApiError: (error: unknown): error is ApiError => error instanceof ApiError,
};
