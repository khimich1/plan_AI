import { act, renderHook } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { createElement, type ReactNode } from "react";
import { describe, expect, it, vi } from "vitest";
import { gsmApi } from "@/features/gsm/api/gsmApi";
import {
  gsmKeys,
  useDownloadGsmUsageReportMutation,
  useGenerateGsmWaybillsMutation,
} from "@/features/gsm/hooks/useGsmQueries";

vi.mock("@/features/gsm/api/gsmApi", () => ({
  gsmApi: {
    generateWaybills: vi.fn(),
    downloadUsageReport: vi.fn(),
  },
}));

vi.mock("@/shared/lib/downloadFile", () => ({
  saveBlobAs: vi.fn(),
}));

const createWrapper = () => {
  const qc = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  const invalidateSpy = vi.spyOn(qc, "invalidateQueries");
  const wrapper = ({ children }: { children: ReactNode }) =>
    createElement(QueryClientProvider, { client: qc }, children);
  return { wrapper, invalidateSpy };
};

describe("useGsmQueries cache invalidation", () => {
  it("invalidates overview after generate so Recalc and tail badges refresh", async () => {
    vi.mocked(gsmApi.generateWaybills).mockResolvedValue({
      waybills: [],
      warnings: [],
      days_created: 1,
      problematic_days: [],
      manual_days: 0,
    });
    const { wrapper, invalidateSpy } = createWrapper();
    const { result } = renderHook(() => useGenerateGsmWaybillsMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({
        vehicle_id: 1,
        period_from: "2026-08-01",
        period_to: "2026-08-31",
      });
    });

    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: gsmKeys.overview(null) });
    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: gsmKeys.waybills(null) });
  });

  it("invalidates overview and waybills after kit download", async () => {
    vi.mocked(gsmApi.downloadUsageReport).mockResolvedValue({
      blob: new Blob(),
      filename: "gsm_usage_report.zip",
    });
    const { wrapper, invalidateSpy } = createWrapper();
    const { result } = renderHook(() => useDownloadGsmUsageReportMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({
        period_from: "2026-07-01",
        period_to: "2026-07-31",
        vehicle_ids: [1],
      });
    });

    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: gsmKeys.overview(null) });
    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: gsmKeys.waybills(null) });
  });
});
