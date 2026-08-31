import { describe, expect, it, vi, beforeEach } from "vitest";
import type { Mock } from "vitest";
import { capacityApi } from "@/features/factory-capacity/api/capacityApi";
import { httpClient } from "@/shared/api/httpClient";
import { ddMmYyyyToIso, maxIsoDate } from "@/features/factory-capacity/lib/dates";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
  },
  resolveApiUrl: (path: string) => path,
}));

const mockGet = httpClient.get as unknown as Mock;

beforeEach(() => {
  vi.clearAllMocks();
});

describe("capacityApi", () => {
  it("GETs capacity-snapshot with target query", async () => {
    mockGet.mockResolvedValue({ status: "green", tracks_needed: 1 });
    await capacityApi.getSnapshot(42, "2026-03-20");
    expect(mockGet).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/42/capacity-snapshot?target=2026-03-20",
    );
  });
});

describe("dates helpers", () => {
  it("converts DD.MM.YYYY to ISO", () => {
    expect(ddMmYyyyToIso("05.06.2026")).toBe("2026-06-05");
    expect(ddMmYyyyToIso("bad")).toBeNull();
  });

  it("picks max ISO date", () => {
    expect(maxIsoDate(["2026-03-01", "2026-05-10", "nope"])).toBe("2026-05-10");
    expect(maxIsoDate([])).toBeNull();
  });
});
