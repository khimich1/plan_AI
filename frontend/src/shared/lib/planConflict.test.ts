import { describe, expect, it, vi, beforeEach } from "vitest";
import type { QueryClient } from "@tanstack/react-query";

import { ApiError } from "@/shared/lib/apiError";
import {
  __testables,
  handlePlanVersionConflict,
  isPlanVersionConflict,
  PLAN_VERSION_CONFLICT_TOAST_MESSAGE,
  showPlanConflictToast,
} from "@/shared/lib/planConflict";

vi.mock("@/features/production/api/productionApi", () => ({
  productionApi: {
    getPlan: vi.fn().mockResolvedValue({ plan_id: "plan-1", version: 4 }),
  },
}));

describe("planConflict", () => {
  beforeEach(() => {
    __testables.resetPlanConflictUiState();
    vi.clearAllMocks();
  });

  it("detects only plan_version_conflict 409", () => {
    expect(
      isPlanVersionConflict(
        new ApiError("conflict", 409, "План", "plan_version_conflict"),
      ),
    ).toBe(true);
    expect(isPlanVersionConflict(new ApiError("busy", 409, "День завершён"))).toBe(
      false,
    );
  });

  it("shows debounced toast message", () => {
    showPlanConflictToast();
    showPlanConflictToast();

    const container = document.getElementById("plan-conflict-toast-root");
    expect(container).not.toBeNull();
    expect(container?.textContent).toBe(PLAN_VERSION_CONFLICT_TOAST_MESSAGE);
    expect(container?.childElementCount).toBe(1);
  });

  it("invalidates production queries and refetches plan on conflict", async () => {
    const invalidateQueries = vi.fn().mockResolvedValue(undefined);
    const fetchQuery = vi.fn().mockResolvedValue({ plan_id: "plan-1" });
    const queryClient = {
      invalidateQueries,
      fetchQuery,
    } as unknown as QueryClient;

    const handled = await handlePlanVersionConflict(
      queryClient,
      new ApiError("conflict", 409, "План", "plan_version_conflict", {
        plan_id: "plan-1",
        expected_version: 3,
      }),
      { showToast: false },
    );

    expect(handled).toBe(true);
    expect(invalidateQueries).toHaveBeenCalledWith({
      queryKey: ["production"],
      refetchType: "active",
    });
    expect(fetchQuery).toHaveBeenCalledWith(
      expect.objectContaining({
        queryKey: ["production", "plan", "plan-1"],
      }),
    );
  });

  it("returns false for non-conflict errors", async () => {
    const queryClient = {
      invalidateQueries: vi.fn(),
      fetchQuery: vi.fn(),
    } as unknown as QueryClient;

    const handled = await handlePlanVersionConflict(
      queryClient,
      new ApiError("busy", 409, "День уже завершён"),
    );

    expect(handled).toBe(false);
    expect(queryClient.invalidateQueries).not.toHaveBeenCalled();
  });
});
