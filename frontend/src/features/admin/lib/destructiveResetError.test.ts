import { describe, expect, it } from "vitest";

import { ApiError } from "@/shared/lib/apiError";
import {
  DESTRUCTIVE_DB_BLOCKED_HINT,
  getDestructiveResetErrorMessage,
  MSG_DESTRUCTIVE_DB_BLOCKED_CLIENT,
} from "./destructiveResetError";

describe("getDestructiveResetErrorMessage", () => {
  it("returns dev hint for blocked destructive reset 403", () => {
    const error = new ApiError(
      "blocked",
      403,
      MSG_DESTRUCTIVE_DB_BLOCKED_CLIENT,
    );

    expect(getDestructiveResetErrorMessage(error)).toBe(DESTRUCTIVE_DB_BLOCKED_HINT);
  });

  it("returns null for other 403 errors", () => {
    const error = new ApiError("forbidden", 403, "Доступ запрещён");

    expect(getDestructiveResetErrorMessage(error)).toBeNull();
  });

  it("returns null for blocked message on non-403 status", () => {
    const error = new ApiError(
      "blocked",
      500,
      MSG_DESTRUCTIVE_DB_BLOCKED_CLIENT,
    );

    expect(getDestructiveResetErrorMessage(error)).toBeNull();
  });

  it("returns null for non-ApiError", () => {
    expect(getDestructiveResetErrorMessage(new Error("fail"))).toBeNull();
  });
});
