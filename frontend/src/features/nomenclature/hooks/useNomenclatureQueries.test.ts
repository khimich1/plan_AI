import { act, renderHook } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { createElement, type ReactNode } from "react";
import { describe, expect, it, vi } from "vitest";
import { guidCheckKeys } from "@/features/commercial-offer/hooks/useGuidCheckQuery";
import { nomenclatureApi } from "@/features/nomenclature/api/nomenclatureApi";
import {
  nomenclatureKeys,
  useImport1cMutation,
} from "@/features/nomenclature/hooks/useNomenclatureQueries";

vi.mock("@/features/nomenclature/api/nomenclatureApi", () => ({
  nomenclatureApi: {
    import1c: vi.fn(),
  },
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

describe("useImport1cMutation", () => {
  it("invalidates nomenclature and guid-check after a successful import", async () => {
    vi.mocked(nomenclatureApi.import1c).mockResolvedValue({
      product_kind: "pile",
      summary: "+1 GUID",
      new_guids_count: 1,
      waiting_price: 0,
      ambiguous_count: 0,
      disappeared_count: 0,
      unmatched_1c_count: 0,
      unchanged: 0,
      updated_guids_count: 0,
      mode: "partial",
      weights_updated: 0,
      list_limit: 50,
      new_guids: [],
      updated_guids: [],
      ambiguous: [],
      disappeared: [],
      unmatched_1c: [],
    });
    const { wrapper, invalidateSpy } = createWrapper();
    const { result } = renderHook(() => useImport1cMutation(), { wrapper });
    const file = new File(["xlsx"], "сваи.xlsx");

    await act(async () => {
      await result.current.mutateAsync({ file, productKind: "pile" });
    });

    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: nomenclatureKeys.all });
    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: guidCheckKeys.all });
  });
});
