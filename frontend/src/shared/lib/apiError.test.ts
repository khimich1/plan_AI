import { describe, expect, it } from "vitest";

import {
  ApiError,
  getErrorMessage,
  getPlanVersionConflictDetails,
  isPlanVersionConflict,
  parseApiErrorPayload,
  PLAN_VERSION_CONFLICT_USER_MESSAGE,
} from "./apiError";

describe("parseApiErrorPayload", () => {
  it("reads structured error nested under detail (FastAPI)", () => {
    const result = parseApiErrorPayload({
      detail: {
        code: "unpriced_plates",
        message: "Нет цен для части позиций",
        details: { positions: ["ПБ 78-12-8п"] },
      },
    });

    expect(result).toEqual({
      message: "Нет цен для части позиций",
      code: "unpriced_plates",
      details: { positions: ["ПБ 78-12-8п"] },
    });
  });

  it("falls back to legacy detail string", () => {
    expect(parseApiErrorPayload({ detail: "Проверьте введённые данные." })).toEqual({
      message: "Проверьте введённые данные.",
    });
  });

  it("reads top-level structured payload", () => {
    expect(
      parseApiErrorPayload({
        code: "plan_version_conflict",
        message: "План был изменён",
      }),
    ).toEqual({
      message: "План был изменён",
      code: "plan_version_conflict",
      details: undefined,
    });
  });

  it("detects plan version conflict ApiError", () => {
    const error = new ApiError(
      "conflict",
      409,
      "План был изменён",
      "plan_version_conflict",
      { plan_id: "plan-1", expected_version: 3 },
    );

    expect(isPlanVersionConflict(error)).toBe(true);
    expect(getPlanVersionConflictDetails(error)).toEqual({
      plan_id: "plan-1",
      expected_version: 3,
    });
    expect(getErrorMessage(error)).toBe(PLAN_VERSION_CONFLICT_USER_MESSAGE);
  });

  it("does not treat generic 409 as plan version conflict", () => {
    const error = new ApiError("busy", 409, "День уже завершён");
    expect(isPlanVersionConflict(error)).toBe(false);
    expect(getErrorMessage(error)).toBe("День уже завершён");
  });
});
