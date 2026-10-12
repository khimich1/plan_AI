import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { pileCostApi } from "@/features/price-desk/api/pileCostApi";
import { httpClient } from "@/shared/api/httpClient";
import type { PileCostInputsBody } from "@/features/price-desk/types/pileCost";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
  },
}));

const mockGet = httpClient.get as unknown as Mock;
const mockPost = httpClient.post as unknown as Mock;
const mockPut = httpClient.put as unknown as Mock;

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

beforeEach(() => {
  vi.clearAllMocks();
});

describe("pileCostApi", () => {
  it("GETs the knob state without a mark", async () => {
    mockGet.mockResolvedValue({ marks: [] });
    await pileCostApi.get();
    expect(mockGet).toHaveBeenCalledWith("/api/v1/prices/pile-cost");
  });

  it("GETs one mark breakdown", async () => {
    mockGet.mockResolvedValue({ breakdown: {} });
    await pileCostApi.get("  С30.20-3  ");
    expect(mockGet).toHaveBeenCalledWith(
      `/api/v1/prices/pile-cost?mark=${encodeURIComponent("С30.20-3")}`,
    );
  });

  it("PUTs knobs and drops volume and rebar kilograms", async () => {
    mockPut.mockResolvedValue({ inputs_sha256: "abc" });
    const body = {
      ...inputs(),
      volume_m3: 0.13,
      lines: [{ article: "a500c_12", kg: 8 }],
    };
    await pileCostApi.save(body);
    const [url, raw, headers] = mockPut.mock.calls[0] as [string, string, HeadersInit];
    expect(url).toBe("/api/v1/prices/pile-cost");
    expect(headers).toEqual({ "Content-Type": "application/json" });
    const sent = JSON.parse(raw) as Record<string, unknown>;
    expect(sent).toEqual(inputs());
    expect(sent).not.toHaveProperty("volume_m3");
    expect(sent).not.toHaveProperty("lines");
    expect(JSON.stringify(sent)).not.toContain("a500c_12");
  });

  it("POSTs preview with no body", async () => {
    mockPost.mockResolvedValue({ inputs_sha256: "abc", changed: 0 });
    await pileCostApi.preview();
    expect(mockPost).toHaveBeenCalledTimes(1);
    expect(mockPost).toHaveBeenCalledWith("/api/v1/prices/pile-cost/preview");
    expect(mockPut).not.toHaveBeenCalled();
  });

  it("POSTs apply with inputs_sha256 only", async () => {
    mockPost.mockResolvedValue({ inputs_sha256: "deadbeef" });
    await pileCostApi.apply({ inputs_sha256: "deadbeef" });
    const [url, raw, headers] = mockPost.mock.calls[0] as [string, string, HeadersInit];
    expect(url).toBe("/api/v1/prices/pile-cost/apply");
    expect(headers).toEqual({ "Content-Type": "application/json" });
    expect(JSON.parse(raw)).toEqual({ inputs_sha256: "deadbeef" });
  });
});
