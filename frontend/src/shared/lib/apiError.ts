export class ApiError extends Error {
  status: number;
  detail: string;
  code?: string;
  details?: unknown;

  constructor(
    message: string,
    status = 500,
    detail = message,
    code?: string,
    details?: unknown,
  ) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.detail = detail;
    this.code = code;
    this.details = details;
  }
}

type StructuredErrorPayload = {
  code: string;
  message: string;
  details?: unknown;
};

const isStructuredErrorPayload = (value: unknown): value is StructuredErrorPayload => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return false;
  }
  const record = value as Record<string, unknown>;
  return typeof record.code === "string" && typeof record.message === "string";
};

export const parseApiErrorPayload = (
  payload: unknown,
): { message: string; code?: string; details?: unknown } | null => {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    return null;
  }
  const record = payload as Record<string, unknown>;

  if (isStructuredErrorPayload(record.detail)) {
    return {
      message: record.detail.message,
      code: record.detail.code,
      details: record.detail.details,
    };
  }

  if (isStructuredErrorPayload(record)) {
    return {
      message: record.message,
      code: record.code,
      details: record.details,
    };
  }

  if (typeof record.detail === "string" && record.detail.trim()) {
    return { message: record.detail };
  }

  return null;
};

export const PLAN_VERSION_CONFLICT_CODE = "plan_version_conflict";

export const PLAN_VERSION_CONFLICT_USER_MESSAGE =
  "План был изменён в другой вкладке или сессии. Данные обновлены — проверьте актуальное состояние и повторите действие при необходимости.";

export type PlanVersionConflictDetails = {
  plan_id?: string;
  expected_version?: number;
};

export const isPlanVersionConflict = (error: unknown): boolean =>
  error instanceof ApiError &&
  error.status === 409 &&
  error.code === PLAN_VERSION_CONFLICT_CODE;

export const getPlanVersionConflictDetails = (
  error: unknown,
): PlanVersionConflictDetails | null => {
  if (!isPlanVersionConflict(error) || !(error instanceof ApiError)) {
    return null;
  }
  const details = error.details;
  if (!details || typeof details !== "object" || Array.isArray(details)) {
    return null;
  }
  const record = details as Record<string, unknown>;
  return {
    plan_id: typeof record.plan_id === "string" ? record.plan_id : undefined,
    expected_version:
      typeof record.expected_version === "number"
        ? record.expected_version
        : undefined,
  };
};

export const getErrorMessage = (error: unknown): string => {
  if (isPlanVersionConflict(error)) {
    return PLAN_VERSION_CONFLICT_USER_MESSAGE;
  }
  if (error instanceof ApiError) {
    return error.detail || error.message;
  }
  if (error instanceof Error) {
    return error.message;
  }
  return "Неизвестная ошибка. Попробуйте ещё раз.";
};
