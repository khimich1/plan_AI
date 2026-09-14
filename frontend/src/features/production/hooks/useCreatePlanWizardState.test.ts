import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type {
  AnalyzeSubstratesResponse,
  CapacityDeficit,
  CapacityOption,
  DayInfo,
  FillTargetItem,
  KpCandidateItem,
  SubstrateRecommendation,
  UrgentPosition,
} from "@/features/production/types/production";
import { useCreatePlanWizardState } from "./useCreatePlanWizardState";

const mockMutate = vi.fn();
const mockAnalyzeMutate = vi.fn();
const mockAnalyzeReset = vi.fn();
const mockUseGlobalCalendarQuery = vi.fn();
const mockUsePlansListQuery = vi.fn();
const mockUseKpCandidatesQuery = vi.fn();
const mockUseBuildPlanMutation = vi.fn();
const mockUseAnalyzeSubstratesMutation = vi.fn();

vi.mock("@/features/production/hooks/useProductionQueries", () => ({
  useGlobalCalendarQuery: () => mockUseGlobalCalendarQuery(),
  usePlansListQuery: () => mockUsePlansListQuery(),
  useKpCandidatesQuery: (enabled: boolean) => mockUseKpCandidatesQuery(enabled),
  useBuildPlanMutation: () => mockUseBuildPlanMutation(),
  useAnalyzeSubstratesMutation: () => mockUseAnalyzeSubstratesMutation(),
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
  mockUseAnalyzeSubstratesMutation.mockReturnValue({
    mutate: mockAnalyzeMutate,
    reset: mockAnalyzeReset,
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

  it("calls analyze-substrates once when fillTargets appear", () => {
    renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    expect(mockAnalyzeMutate).toHaveBeenCalledWith(
      {
        fill_targets: partialFillRequest,
        deadline_until: "2026-06-22",
      },
      expect.any(Object),
    );
  });

  it("analyze success preselects urgent plates (default on) into selection+qty", () => {
    const urgent: UrgentPosition = {
      plate_id: 123,
      kp_id: 115,
      plate_name: "ПБ 57-7,2 ×8п",
      qty_remaining: 2,
      deadline: "2026-06-22",
      deadline_source: "delivery_batch",
      deadline_details: [],
      conflict: null,
    };
    const analyzeData: AnalyzeSubstratesResponse = {
      urgent_positions: [urgent],
      substrate_recommendations: [],
      capacity_deficit: null,
      analysis_meta: {
        orders_count: 1,
        analysis_duration_ms: 10,
        optimization_status: "ok",
        error_message: null,
      },
    };

    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    const onSuccess = mockAnalyzeMutate.mock.calls[0][1]?.onSuccess as
      | ((data: AnalyzeSubstratesResponse) => void)
      | undefined;
    expect(onSuccess).toBeTypeOf("function");

    act(() => {
      onSuccess!(analyzeData);
    });

    expect(result.current.analyzeResult?.urgent_positions).toEqual([urgent]);
    expect(result.current.selectedPlatesByKp).toEqual({ 115: [123] });
    expect(result.current.selectedPlateQtyByKp).toEqual({ 115: { 123: 2 } });
    expect(result.current.filterMethod).toBe("kp");
  });

  it("toggleUrgentPosition syncs selectedPlatesByKp and selectedPlateQtyByKp", () => {
    const urgent: UrgentPosition = {
      plate_id: 123,
      kp_id: 115,
      plate_name: "ПБ 57-7,2 ×8п",
      qty_remaining: 2,
      deadline: "2026-06-22",
      deadline_source: "delivery_batch",
      deadline_details: [],
      conflict: null,
    };
    const analyzeData: AnalyzeSubstratesResponse = {
      urgent_positions: [urgent],
      substrate_recommendations: [],
      capacity_deficit: null,
      analysis_meta: {
        orders_count: 1,
        analysis_duration_ms: 10,
        optimization_status: "ok",
        error_message: null,
      },
    };

    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    const onSuccess = mockAnalyzeMutate.mock.calls[0][1]?.onSuccess as
      | ((data: AnalyzeSubstratesResponse) => void)
      | undefined;

    act(() => {
      onSuccess!(analyzeData);
    });

    act(() => {
      result.current.toggleUrgentPosition(urgent);
    });

    expect(result.current.selectedPlatesByKp).toEqual({});
    expect(result.current.selectedPlateQtyByKp).toEqual({});

    act(() => {
      result.current.toggleUrgentPosition(urgent);
    });

    expect(result.current.selectedPlatesByKp).toEqual({ 115: [123] });
    expect(result.current.selectedPlateQtyByKp).toEqual({ 115: { 123: 2 } });
  });

  it("analyze success preselects substrate recommendations with qty_recommended", () => {
    const substrate: SubstrateRecommendation = {
      plate_id: 456,
      kp_id: 127,
      plate_name: "ПБ 57-4,8 ×8п",
      qty_recommended: 3,
      under_plate_id: 123,
      under_kp_id: 115,
      under_plate_name: "ПБ 57-7,2 ×8п",
      needed_by: "2026-09-05",
      storage_days: 24,
      saving_mm: 480,
      saving_m: 2.4,
    };
    const analyzeData: AnalyzeSubstratesResponse = {
      urgent_positions: [],
      substrate_recommendations: [substrate],
      capacity_deficit: null,
      analysis_meta: {
        orders_count: 1,
        analysis_duration_ms: 10,
        optimization_status: "ok",
        error_message: null,
      },
    };

    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    const onSuccess = mockAnalyzeMutate.mock.calls[0][1]?.onSuccess as
      | ((data: AnalyzeSubstratesResponse) => void)
      | undefined;

    act(() => {
      onSuccess!(analyzeData);
    });

    expect(result.current.analyzeResult?.substrate_recommendations).toEqual([
      substrate,
    ]);
    expect(result.current.selectedPlatesByKp).toEqual({ 127: [456] });
    expect(result.current.selectedPlateQtyByKp).toEqual({ 127: { 456: 3 } });
    expect(result.current.filterMethod).toBe("kp");
  });

  it("runAnalyzeSubstrates re-triggers mutation; toggleSubstrateRecommendation syncs selection", () => {
    const substrate: SubstrateRecommendation = {
      plate_id: 456,
      kp_id: 127,
      plate_name: "ПБ 57-4,8 ×8п",
      qty_recommended: 3,
      under_plate_id: 123,
      under_kp_id: 115,
      under_plate_name: "ПБ 57-7,2 ×8п",
      needed_by: "2026-09-05",
      storage_days: 24,
      saving_mm: 480,
      saving_m: 2.4,
    };
    const analyzeData: AnalyzeSubstratesResponse = {
      urgent_positions: [],
      substrate_recommendations: [substrate],
      capacity_deficit: null,
      analysis_meta: {
        orders_count: 1,
        analysis_duration_ms: 10,
        optimization_status: "ok",
        error_message: null,
      },
    };

    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    expect(mockAnalyzeMutate).toHaveBeenCalledTimes(1);

    act(() => {
      result.current.runAnalyzeSubstrates();
    });

    expect(mockAnalyzeMutate).toHaveBeenCalledTimes(2);

    const onSuccess = mockAnalyzeMutate.mock.calls[1][1]?.onSuccess as
      | ((data: AnalyzeSubstratesResponse) => void)
      | undefined;

    act(() => {
      onSuccess!(analyzeData);
    });

    act(() => {
      result.current.toggleSubstrateRecommendation(substrate);
    });

    expect(result.current.selectedPlatesByKp).toEqual({});
    expect(result.current.selectedPlateQtyByKp).toEqual({});

    act(() => {
      result.current.toggleSubstrateRecommendation(substrate);
    });

    expect(result.current.selectedPlatesByKp).toEqual({ 127: [456] });
    expect(result.current.selectedPlateQtyByKp).toEqual({ 127: { 456: 3 } });
  });

  it("applyCapacityOption bumps fillTargets without saving day capacity", () => {
    const deficit: CapacityDeficit = {
      tracks_needed: 10,
      tracks_available: 5,
      tracks_missing: 2,
      deficit_until: "2026-06-22",
      options: [
        { action: "bump_fill", date: "2026-06-21", add_tracks: 2, free: 5 },
      ],
    };
    const analyzeData: AnalyzeSubstratesResponse = {
      urgent_positions: [],
      substrate_recommendations: [],
      capacity_deficit: deficit,
      analysis_meta: {
        orders_count: 1,
        analysis_duration_ms: 10,
        optimization_status: "ok",
        error_message: null,
      },
    };

    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    const onSuccess = mockAnalyzeMutate.mock.calls[0][1]?.onSuccess as
      | ((data: AnalyzeSubstratesResponse) => void)
      | undefined;

    act(() => {
      onSuccess!(analyzeData);
    });

    expect(mockAnalyzeMutate).toHaveBeenCalledTimes(1);

    const option: CapacityOption = deficit.options[0];
    act(() => {
      result.current.applyCapacityOption(option);
    });

    expect(result.current.fillTargets).toEqual([
      { date: "2026-06-21", tracks: 4 },
      { date: "2026-06-22", tracks: 3 },
    ]);
    // targetsKey includes tracks → analyze re-runs
    expect(mockAnalyzeMutate).toHaveBeenCalledTimes(2);
    expect(mockAnalyzeMutate).toHaveBeenLastCalledWith(
      {
        fill_targets: [
          { date: "2026-06-21", tracks: 4 },
          { date: "2026-06-22", tracks: 3 },
        ],
        deadline_until: "2026-06-22",
      },
      expect.any(Object),
    );
  });

  it("applyCapacityOption propose_day adds a new date to fillTargets", () => {
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    act(() => {
      result.current.applyCapacityOption({
        action: "propose_day",
        date: "2026-06-20",
        add_tracks: 3,
        free: 3,
      });
    });

    expect(result.current.fillTargets).toEqual([
      { date: "2026-06-20", tracks: 3 },
      { date: "2026-06-21", tracks: 2 },
      { date: "2026-06-22", tracks: 3 },
    ]);
  });

  it("exposes substrateErrorMessage from analysis_meta when optimization_status is error", () => {
    const analyzeData: AnalyzeSubstratesResponse = {
      urgent_positions: [],
      substrate_recommendations: [],
      capacity_deficit: null,
      analysis_meta: {
        orders_count: 2,
        analysis_duration_ms: 12,
        optimization_status: "error",
        error_message:
          "Оптимизатор вернул ошибку при анализе подложек: infeasible",
      },
    };

    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      {
        initialProps: { fillRequest: partialFillRequest },
      },
    );

    const onSuccess = mockAnalyzeMutate.mock.calls[0][1]?.onSuccess as
      | ((data: AnalyzeSubstratesResponse) => void)
      | undefined;

    act(() => {
      onSuccess!(analyzeData);
    });

    expect(result.current.analyzeErrorMessage).toBeNull();
    expect(result.current.substrateErrorMessage).toBe(
      "Оптимизатор вернул ошибку при анализе подложек: infeasible",
    );
  });
});

const promisedPlate = {
  id: 501,
  plate_name: "ПБ 60-12-8п",
  length_m: 6,
  width_m: 1.2,
  load_class: 800,
  qty: 2,
};

function makePromisedKp(overrides: Partial<KpCandidateItem> = {}): KpCandidateItem {
  return {
    kp_id: 88,
    customer_name: "Обещанный",
    creation_date: "01.09.2026",
    execution_terms: "25.09.2026",
    total_plates: 1,
    completed_plates: 0,
    completion_pct: 0,
    in_plan_pct: 0,
    total_length_m: 12,
    plates: [promisedPlate],
    promise: {
      promised_date: "2026-09-25",
      week_start: "2026-06-15",
      status: "active",
      tracks: 2,
    },
    ...overrides,
  };
}

function mockPromisedCandidates(items: KpCandidateItem[]) {
  mockUseKpCandidatesQuery.mockReturnValue({
    data: {
      items,
      count: items.length,
      promised_weeks: items
        .filter((kp) => kp.promise)
        .map((kp) => ({
          week_start: kp.promise!.week_start,
          items: [
            {
              kp_id: kp.kp_id,
              promised_date: kp.promise!.promised_date,
              tracks: kp.promise!.tracks,
              status: kp.promise!.status,
            },
          ],
        })),
    },
    isLoading: false,
  });
}

describe("useCreatePlanWizardState promised KPs", () => {
  beforeEach(() => {
    setupQueryMocks();
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("preselects promised KP plates for the fill week", () => {
    mockPromisedCandidates([makePromisedKp()]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: partialFillRequest } },
    );

    expect(result.current.filterMethod).toBe("kp");
    expect(result.current.selectedPlatesByKp).toEqual({ 88: [501] });
    expect(result.current.selectedPlateQtyByKp).toEqual({ 88: { 501: 2 } });
    expect(result.current.promisedBlockItems.map((item) => item.kp_id)).toEqual([88]);
  });

  it("preselects overdue promised KPs even outside the fill week", () => {
    mockPromisedCandidates([
      makePromisedKp({
        kp_id: 91,
        customer_name: "Просроченный",
        promise: {
          promised_date: "2026-09-04",
          week_start: "2026-06-08",
          status: "overdue",
          tracks: 3,
        },
      }),
    ]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: partialFillRequest } },
    );

    expect(result.current.selectedPlatesByKp).toEqual({ 91: [501] });
    expect(result.current.promisedBlockItems[0]?.status).toBe("overdue");
  });

  it("does not uncheck a promised KP without a reason", () => {
    const kp = makePromisedKp();
    mockPromisedCandidates([kp]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: partialFillRequest } },
    );

    act(() => {
      result.current.toggleKp(kp);
    });

    expect(result.current.selectedPlatesByKp).toEqual({ 88: [501] });
    expect(result.current.pendingExclusion).toEqual({
      kpId: 88,
      weekStart: "2026-06-15",
      kind: "whole",
    });
    expect(result.current.canSubmit).toBe(false);
  });

  it("rejects a blank exclusion reason and keeps the KP selected", () => {
    const kp = makePromisedKp();
    mockPromisedCandidates([kp]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: partialFillRequest } },
    );

    act(() => {
      result.current.toggleKp(kp);
    });
    act(() => {
      result.current.confirmExclusion("   ");
    });

    expect(result.current.selectedPlatesByKp).toEqual({ 88: [501] });
    expect(result.current.exclusions).toEqual([]);
    expect(result.current.pendingExclusion).not.toBeNull();
  });

  it("stores the reason, deselects the KP, and sends exclusions on submit", () => {
    const kp = makePromisedKp();
    mockPromisedCandidates([kp]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: partialFillRequest } },
    );

    act(() => {
      result.current.toggleKp(kp);
    });
    act(() => {
      result.current.confirmExclusion("Клиент перенёс поставку");
    });

    expect(result.current.selectedPlatesByKp).toEqual({});
    expect(result.current.pendingExclusion).toBeNull();
    expect(result.current.exclusions).toEqual([
      {
        kp_id: 88,
        week_start: "2026-06-15",
        reason: "Клиент перенёс поставку",
      },
    ]);

    act(() => {
      result.current.handleSubmit("asc");
    });

    expect(mockMutate).toHaveBeenCalledTimes(1);
    expect(mockMutate.mock.calls[0][0].exclusions).toEqual([
      {
        kp_id: 88,
        week_start: "2026-06-15",
        reason: "Клиент перенёс поставку",
      },
    ]);
  });

  it("requires a reason to uncheck a plate of a promised KP", () => {
    const kp = makePromisedKp({
      plates: [
        promisedPlate,
        {
          id: 502,
          plate_name: "ПБ 51-12-8п",
          length_m: 5.1,
          width_m: 1.2,
          load_class: 800,
          qty: 1,
        },
      ],
    });
    mockPromisedCandidates([kp]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: partialFillRequest } },
    );

    act(() => {
      result.current.togglePlate(kp, 502);
    });

    expect(result.current.selectedPlatesByKp[88]).toEqual([501, 502]);
    expect(result.current.pendingExclusion).toEqual({
      kpId: 88,
      weekStart: "2026-06-15",
      kind: "partial",
      plateId: 502,
    });

    act(() => {
      result.current.confirmExclusion("Часть позиций ушла в другой план");
    });

    expect(result.current.selectedPlatesByKp[88]).toEqual([501]);
    expect(result.current.exclusions[0]?.reason).toBe(
      "Часть позиций ушла в другой план",
    );
  });

  it("omits exclusions from the build payload when none were recorded", () => {
    mockPromisedCandidates([makePromisedKp()]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: emptyFillRequest } },
    );

    act(() => {
      result.current.handleSubmit("asc");
    });

    expect(mockMutate.mock.calls[0][0].exclusions).toBeUndefined();
  });

  it("clears the exclusion when the promised KP is selected again", () => {
    const kp = makePromisedKp();
    mockPromisedCandidates([kp]);
    const { result } = renderHook(
      (props: Parameters<typeof useCreatePlanWizardState>[0]) =>
        useCreatePlanWizardState(props),
      { initialProps: { fillRequest: partialFillRequest } },
    );

    act(() => {
      result.current.toggleKp(kp);
      result.current.confirmExclusion("Временно снимаем");
    });
    act(() => {
      result.current.toggleKp(kp);
    });

    expect(result.current.selectedPlatesByKp).toEqual({ 88: [501] });
    expect(result.current.exclusions).toEqual([]);
  });
});


