import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { nomenclatureApi } from "@/features/nomenclature/api/nomenclatureApi";
import { httpClient } from "@/shared/api/httpClient";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    post: vi.fn(),
    get: vi.fn(),
  },
}));

const mockPost = httpClient.post as unknown as Mock;
const mockGet = httpClient.get as unknown as Mock;

beforeEach(() => {
  vi.clearAllMocks();
});

describe("nomenclatureApi.import1c", () => {
  const file = new File(["xls"], "Прайс сваи.xls", { type: "application/vnd.ms-excel" });

  it("POSTs FormData file without product_kind when omitted", async () => {
    mockPost.mockResolvedValue({ summary: "+0 GUID" });

    await nomenclatureApi.import1c(file);

    expect(mockPost).toHaveBeenCalledTimes(1);
    const [url, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(url).toBe("/api/v1/nomenclature/import-1c");
    expect(body).toBeInstanceOf(FormData);
    expect(body.get("file")).toBe(file);
    expect(body.has("product_kind")).toBe(false);
  });

  it("appends product_kind for mixed ЛМ dumps", async () => {
    mockPost.mockResolvedValue({ summary: "+0 GUID" });

    await nomenclatureApi.import1c(file, "stair_flight");

    const [, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(body.get("product_kind")).toBe("stair_flight");
  });
});

describe("nomenclatureApi queues", () => {
  it("GETs /tasks", async () => {
    mockGet.mockResolvedValue({ to_create_1c: [], to_price: [], duplicates: [] });
    await nomenclatureApi.listTasks();
    expect(mockGet).toHaveBeenCalledWith("/api/v1/nomenclature/tasks");
  });

  it("POSTs price-queue resolve", async () => {
    mockPost.mockResolvedValue({ guid: "aaa", mark: "С30.30-3", product_kind: "pile", price: 1500, state: "resolved" });
    await nomenclatureApi.resolvePrice("aaa", 1500);
    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/nomenclature/price-queue/resolve",
      JSON.stringify({ guid: "aaa", price: 1500 }),
      { "Content-Type": "application/json" },
    );
  });

  it("POSTs duplicates resolve", async () => {
    mockPost.mockResolvedValue({ scope: "nonplate", key: "С40.30-6", chosen_guid: "bbb", match_status: "manual" });
    await nomenclatureApi.resolveDuplicate({
      scope: "nonplate",
      key: "С40.30-6",
      chosen_guid: "bbb",
      note: "бухгалтерия",
    });
    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/nomenclature/duplicates/resolve",
      JSON.stringify({
        scope: "nonplate",
        key: "С40.30-6",
        chosen_guid: "bbb",
        note: "бухгалтерия",
        product_kind: null,
      }),
      { "Content-Type": "application/json" },
    );
  });
});
