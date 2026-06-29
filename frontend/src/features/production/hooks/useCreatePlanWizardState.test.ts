import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { FillTargetItem } from "@/features/production/types/production";
import { useCreatePlanWizardState } from "./useCreatePlanWizardState";

const mockMutate = vi.fn();
const mockUseDayOccupancyQuery = vi.fn();
const mockUseGlobalCalendarQuery = vi.fn();
const mockUsePlansListQuery = vi.fn();
const mockUseWorkCalendarQuery = vi.fn();
const mockUseKpCandidatesQuery = vi.fn();
const mockUseBuildPlanMutation = vi.fn();

vi.mock("@/features/production/hooks/useProductionQueries", () => ({
  useDayOccupancyQuery: () => mockUseDayOccupancyQuery(),
  useGlobalCalendarQuery: () => mockUseGlobalCalendarQuery(),
  usePlansListQuery: () => mockUsePlansListQuery(),
  useWorkCalendarQuery: () => mockUseWorkCalendarQuery(),
  useKpCandidatesQuery: (enabled: boolean) => mockUseKpCandidatesQuery(enabled),
  useBuildPlanMutation: () => mockUseBuildPlanMutation(),
}));

function setupQueryMocks() {
  mockUseDayOccupancyQuery.mockReturnValue({
    data: { occupancy: { "2026-06-21": 1 }, max_per_day: 5 },
    isLoading: false,
  });
  mockUseGlobalCalendarQuery.mockReturnValue({
    data: { days_info: {} },
    isLoading: false,
  });
  mockUsePlansListQuery.mockReturnValue({
    data: { plans: [], active_plan_id: null },
    isLoading: false,
  });
  mockUseWorkCalendarQuery.mockReturnValue({
    data: { extra_holidays: [], extra_workdays: [] },
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

const fillRequest: FillTargetItem[] = [
  { date: "2026-06-21", tracks: 2 },
  { date: "2026-06-22", tracks: 3 },
];

describe("useCreatePlanWizardState", () => {
  beforeEach(() => {
    setupQueryMocks();
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("starts on step 1 with normal wizard title", () => {
    const { result } = renderHook(() => useCreatePlanWizardState({}));

    expect(result.current.step).toBe(1);
    expect(result.current.isFillMode).toBe(false);
    expect(result.current.cardTitle).toBe("Начать планирование");
    expect(mockUseKpCandidatesQuery).toHaveBeenCalledWith(false);
  });

  it("advances and retreats through wizard steps", () => {
    const { result } = renderHook(() => useCreatePlanWizardState({}));

    act(() => {
      result.current.setStep(2);
    });
    expect(result.current.step).toBe(2);
    expect(mockUseKpCandidatesQuery).toHaveBeenLastCalledWith(false);

    act(() => {
      result.current.setStep(3);
    });
    expect(result.current.step).toBe(3);
    expect(mockUseKpCandidatesQuery).toHaveBeenLastCalledWith(true);

    act(() => {
      result.current.setStep(1);
    });
    expect(result.current.step).toBe(1);
  });

  it("validates step 2 track count bounds", () => {
    const { result } = renderHook(() => useCreatePlanWizardState({}));

    expect(result.current.canProceedStep2).toBe(true);

    act(() => {
      result.current.setTracksCount(0);
    });
    expect(result.current.canProceedStep2).toBe(false);

    act(() => {
      result.current.setTracksCount(51);
    });
    expect(result.current.canProceedStep2).toBe(false);

    act(() => {
      result.current.setTracksCount(10);
    });
    expect(result.current.canProceedStep2).toBe(true);
  });

  it("enters fill mode from fillRequest and calls onFillRequestConsumed", () => {
    const onFillRequestConsumed = vi.fn();
    const { result, rerender } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { onFillRequestConsumed },
      },
    );

    expect(result.current.step).toBe(1);
    expect(result.current.isFillMode).toBe(false);

    rerender({ fillRequest, onFillRequestConsumed });

    expect(result.current.step).toBe(3);
    expect(result.current.isFillMode).toBe(true);
    expect(result.current.cardTitle).toBe("Дозаполнение дней");
    expect(result.current.cardSubtitle).toContain("5 дор.");
    expect(result.current.tracksPerDay).toBe(3);
    expect(result.current.tracksPerDaySource).toBe("дозаполнение");
    expect(onFillRequestConsumed).toHaveBeenCalledTimes(1);
    expect(mockUseKpCandidatesQuery).toHaveBeenLastCalledWith(true);
  });

  it("handleCancelFill exits fill mode and resets selection", () => {
    const onCancelFill = vi.fn();
    const { result, rerender } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest, onCancelFill },
      },
    );

    expect(result.current.isFillMode).toBe(true);
    expect(result.current.step).toBe(3);

    act(() => {
      result.current.handleCancelFill();
    });

    expect(result.current.isFillMode).toBe(false);
    expect(result.current.step).toBe(1);
    expect(result.current.selectedPlatesByKp).toEqual({});
    expect(onCancelFill).toHaveBeenCalledTimes(1);

    rerender({ fillRequest: null, onCancelFill });
    expect(result.current.isFillMode).toBe(false);
  });

  it("canSubmit is true in fill mode without step 1/2 validation", () => {
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest },
      },
    );

    act(() => {
      result.current.setTracksCount(0);
    });

    expect(result.current.canProceedStep2).toBe(false);
    expect(result.current.canSubmit).toBe(true);

    act(() => {
      result.current.handleCancelFill();
    });
    expect(result.current.canSubmit).toBe(false);
  });
});
