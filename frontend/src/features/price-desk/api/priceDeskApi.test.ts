import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { priceDeskApi } from "@/features/price-desk/api/priceDeskApi";
import { httpClient } from "@/shared/api/httpClient";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    get: vi.fn(),
    post: vi.fn(),
  },
}));

const mockGet = httpClient.get as unknown as Mock;
const mockPost = httpClient.post as unknown as Mock;

beforeEach(() => {
  vi.clearAllMocks();
});

describe("priceDeskApi", () => {
  const file = new File(["xlsx"], "Прайс ЛМ от 07.09.2026.xlsx", {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  it("GETs status", async () => {
    mockGet.mockResolvedValue({ groups: [] });
    await priceDeskApi.status();
    expect(mockGet).toHaveBeenCalledWith("/api/v1/prices/status");
  });

  it("POSTs preview FormData with file", async () => {
    mockPost.mockResolvedValue({ file_sha256: "abc" });
    await priceDeskApi.preview(file);
    const [url, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(url).toBe("/api/v1/prices/preview");
    expect(body.get("file")).toBe(file);
  });

  it("POSTs apply FormData with file and sha256", async () => {
    mockPost.mockResolvedValue({ file_sha256: "deadbeef" });
    await priceDeskApi.apply(file, "deadbeef");
    const [url, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(url).toBe("/api/v1/prices/apply");
    expect(body.get("file")).toBe(file);
    expect(body.get("file_sha256")).toBe("deadbeef");
  });
});
