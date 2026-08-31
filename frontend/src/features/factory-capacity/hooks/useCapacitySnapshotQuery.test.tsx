import { describe, expect, it } from "vitest";
import { renderHook } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { useCapacitySnapshotQuery } from "@/features/factory-capacity/hooks/useCapacitySnapshotQuery";

const wrapper = ({ children }: { children: ReactNode }) => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false } },
  });
  return <QueryClientProvider client={client}>{children}</QueryClientProvider>;
};

describe("useCapacitySnapshotQuery", () => {
  it("is disabled when kpId is null", () => {
    const { result } = renderHook(
      () => useCapacitySnapshotQuery(null, "2026-03-20"),
      { wrapper },
    );
    expect(result.current.fetchStatus).toBe("idle");
    expect(result.current.isFetching).toBe(false);
  });

  it("is disabled when target is null", () => {
    const { result } = renderHook(
      () => useCapacitySnapshotQuery(42, null),
      { wrapper },
    );
    expect(result.current.fetchStatus).toBe("idle");
  });
});
