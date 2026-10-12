import { act, renderHook, waitFor } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { createElement, type ReactNode } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { pileCostApi } from "@/features/price-desk/api/pileCostApi";
import { priceDeskKeys } from "@/features/price-desk/hooks/usePriceDeskQueries";
import {
  pileCostKeys,
  usePileCostApplyMutation,
  usePileCostPreviewMutation,
  usePileCostQuery,
  useSavePileCostInputsMutation,
} from "@/features/price-desk/hooks/usePileCostQueries";
import type { PileCostInputsBody, PileCostState } from "@/features/price-desk/types/pileCost";

vi.mock("@/features/price-desk/api/pileCostApi", () => ({
  pileCostApi: {
    get: vi.fn(),
    save: vi.fn(),
    preview: vi.fn(),
    apply: vi.fn(),
  },
}));

const inputs = (): PileCostInputsBody => ({
  prices: [{ article: "cement", price: 7.54 }],
  recipes: [
    {
      concrete_grade: "B15",
      cement_kg_m3: 320,
      sand_kg_m3: 990,
      crushed_kg_m3: 950,
    },
  ],
  overhead: { general: 3186, payroll: 3400, forming: 40, other: 606 },
});

const state = (): PileCostState => ({
  inputs_sha256: "abc",
  inputs: inputs(),
  marks: [],
  breakdown: null,
});

const createWrapper = () => {
  const qc = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  const invalidateSpy = vi.spyOn(qc, "invalidateQueries");
  const wrapper = ({ children }: { children: ReactNode }) =>
    createElement(QueryClientProvider, { client: qc }, children);
  return { wrapper, invalidateSpy };
};

beforeEach(() => {
  vi.clearAllMocks();
});

describe("pile cost hooks", () => {
  it("reads a mark and does not preview or apply", async () => {
    vi.mocked(pileCostApi.get).mockResolvedValue(state());
    const { wrapper } = createWrapper();
    const { result } = renderHook(() => usePileCostQuery("С30.20-3"), { wrapper });

    await waitFor(() => expect(result.current.isSuccess).toBe(true));

    expect(pileCostApi.get).toHaveBeenCalledWith("С30.20-3");
    expect(pileCostApi.preview).not.toHaveBeenCalled();
    expect(pileCostApi.apply).not.toHaveBeenCalled();
  });

  it("saves knobs without touching pile price status", async () => {
    vi.mocked(pileCostApi.save).mockResolvedValue({
      inputs_sha256: "next",
      inputs: inputs(),
    });
    const { wrapper, invalidateSpy } = createWrapper();
    const { result } = renderHook(() => useSavePileCostInputsMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync(inputs());
    });

    expect(pileCostApi.save).toHaveBeenCalledWith(inputs());
    expect(pileCostApi.apply).not.toHaveBeenCalled();
    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: pileCostKeys.all });
    expect(invalidateSpy).not.toHaveBeenCalledWith({ queryKey: priceDeskKeys.status });
  });

  it("previews without writing or invalidating", async () => {
    vi.mocked(pileCostApi.preview).mockResolvedValue({
      inputs_sha256: "abc",
      changed: 1,
      new: 0,
      missing: 0,
      unchanged: 0,
      examples: { changed: [], new: [], missing: [], unchanged: [] },
    });
    const { wrapper, invalidateSpy } = createWrapper();
    const { result } = renderHook(() => usePileCostPreviewMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync();
    });

    expect(pileCostApi.preview).toHaveBeenCalledTimes(1);
    expect(pileCostApi.apply).not.toHaveBeenCalled();
    expect(invalidateSpy).not.toHaveBeenCalled();
  });

  it("applies only the sha256 the caller passes", async () => {
    vi.mocked(pileCostApi.apply).mockResolvedValue({
      inputs_sha256: "deadbeef",
      changed: 0,
      new: 5,
      missing: 0,
      unchanged: 0,
      examples: { changed: [], new: [], missing: [], unchanged: [] },
    });
    const { wrapper, invalidateSpy } = createWrapper();
    const { result } = renderHook(() => usePileCostApplyMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync("deadbeef");
    });

    expect(pileCostApi.apply).toHaveBeenCalledWith({ inputs_sha256: "deadbeef" });
    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: pileCostKeys.all });
    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: priceDeskKeys.status });
  });
});
