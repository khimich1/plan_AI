import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { DayInfo, FillTargetItem } from "@/features/production/types/production";
import { useCreatePlanWizardState } from "./useCreatePlanWizardState";

const mockMutate = vi.fn();
const mockUseGlobalCalendarQuery = vi.fn();
const mockUsePlansListQuery = vi.fn();
const mockUseKpCandidatesQuery = vi.fn();
const mockUseBuildPlanMutation = vi.fn();

vi.mock("@/features/production/hooks/useProductionQueries", () => ({
  useGlobalCalendarQuery: () => mockUseGlobalCalendarQuery(),
  usePlansListQuery: () => mockUsePlansListQuery(),
  useKpCandidatesQuery: (enabled: boolean) => mockUseKpCandidatesQuery(enabled),
  useBuildPlanMutation: () => mockUseBuildPlanMutation(),
}));

vi.mock("@/features/production/hooks/useSgpQueries", () => ({
  useSgpFreePlatesQuery: () => ({
    data: { items: [], count: 0 },
    isLoading: false,
  }),
}));

const daysInfo: Record<string, DayInfo> = {
  "2026-06-21": {
    occupied: 2,
    max: 5,
    completed: false,
    day_number: 1,
  },
  "2026-06-22": {
    occupied: 1,
    max: 5,
    completed: false,
    day_number: 2,
  },
  "2026-07-23": {
    occupied: 0,
    max: 5,
    completed: false,
    day_number: 3,
  },
  "2026-07-25": {
    occupied: 0,
    max: 5,
    completed: false,
    day_number: 4,
  },
};

function setupQueryMocks() {
  mockUseGlobalCalendarQuery.mockReturnValue({
    data: { days_info: daysInfo },
    isLoading: false,
  });
  mockUsePlansListQuery.mockReturnValue({
    data: { plans: [], active_plan_id: null },
    isLoading: false,
  });
  mockUseKpCandidatesQuery.mockReturnValue({
    data: { items: [] },
    isLoading: false,
  });
  mockUseBuildPlanMutation.mockReturnValue({
    mutate: mockMutate,
    isPending: false,
    isSuccess: false,
    isError: false,
    error: null,
  });
}

const partialFillRequest: FillTargetItem[] = [
  { date: "2026-06-21", tracks: 2 },
  { date: "2026-06-22", tracks: 3 },
];

const emptyFillRequest: FillTargetItem[] = [
  { date: "2026-07-23", tracks: 3 },
  { date: "2026-07-25", tracks: 5 },
];

describe("useCreatePlanWizardState", () => {
  beforeEach(() => {
    setupQueryMocks();
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("without fillRequest is not in fill mode and cannot submit", () => {
    const { result } = renderHook(() => useCreatePlanWizardState({}));

    expect(result.current.isFillMode).toBe(false);
    expect(result.current.canSubmit).toBe(false);
    expect(mockUseKpCandidatesQuery).toHaveBeenCalledWith(true);
  });

  it("enters fill mode from fillRequest with auto planName and partial title", () => {
    const onFillRequestConsumed = vi.fn();
    const { result, rerender } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { onFillRequestConsumed },
      },
    );

    expect(result.current.isFillMode).toBe(false);

    rerender({ fillRequest: partialFillRequest, onFillRequestConsumed });

    expect(result.current.isFillMode).toBe(true);
    expect(result.current.cardTitle).toBe("Дозаполнение дней");
    expect(result.current.planName).toBe("План 21–22.06");
    expect(result.current.cardSubtitle).toContain("5 дор.");
    expect(result.current.tracksPerDay).toBe(3);
    expect(result.current.tracksPerDaySource).toBe("календарь");
    expect(result.current.basketKind).toBe("partial");
    expect(onFillRequestConsumed).toHaveBeenCalledTimes(1);
  });

  it("uses empty-kind title and auto planName for free days", () => {
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: emptyFillRequest },
      },
    );

    expect(result.current.isFillMode).toBe(true);
    expect(result.current.cardTitle).toBe("Начать планирование");
    expect(result.current.planName).toBe("План 23–25.07");
    expect(result.current.basketKind).toBe("empty");
  });

  it("handleCancelFill exits fill mode and resets selection", () => {
    const onCancelFill = vi.fn();
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest, onCancelFill },
      },
    );

    expect(result.current.isFillMode).toBe(true);

    act(() => {
      result.current.handleCancelFill();
    });

    expect(result.current.isFillMode).toBe(false);
    expect(result.current.selectedPlatesByKp).toEqual({});
    expect(result.current.planName).toBe("");
    expect(onCancelFill).toHaveBeenCalledTimes(1);
  });

  it("handleSubmit always sends plan_name and fill_targets", () => {
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: emptyFillRequest },
      },
    );

    act(() => {
      result.current.handleSubmit("asc");
    });

    expect(mockMutate).toHaveBeenCalledTimes(1);
    const payload = mockMutate.mock.calls[0][0];
    expect(payload.plan_name).toBe("План 23–25.07");
    expect(payload.fill_targets).toEqual(emptyFillRequest);
    expect(payload.start_date).toBe("2026-07-23");
    expect(payload.tracks_count).toBe(5);
  });

  it("canSubmit is true in fill mode with filterMethod all", () => {
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    expect(result.current.canSubmit).toBe(true);
  });
});
