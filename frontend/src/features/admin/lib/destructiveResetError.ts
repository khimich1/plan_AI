import { ApiError } from "@/shared/lib/apiError";

export const MSG_DESTRUCTIVE_DB_BLOCKED_CLIENT =
  "Операция обнуления базы данных запрещена в текущем окружении.";

export const DESTRUCTIVE_DB_BLOCKED_HINT =
  "Операция запрещена в текущем окружении. Для локальной разработки добавьте " +
  "ALLOW_DESTRUCTIVE_DB_RESET=1 в .env и перезапустите backend.";

export function getDestructiveResetErrorMessage(error: unknown): string | null {
  if (!(error instanceof ApiError) || error.status !== 403) {
    return null;
  }
  if (error.detail !== MSG_DESTRUCTIVE_DB_BLOCKED_CLIENT) {
    return null;
  }
  return DESTRUCTIVE_DB_BLOCKED_HINT;
}
