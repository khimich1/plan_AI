import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { httpClient } from "@/shared/api/httpClient";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
    download: vi.fn(),
  },
  resolveApiUrl: (path: string) => path,
}));

const mockPost = httpClient.post as unknown as Mock;
const mockDownload = httpClient.download as unknown as Mock;

beforeEach(() => {
  vi.clearAllMocks();
});

describe("archiveApi MNA-602 — resume", () => {
  it("exposes resume helper", () => {
    expect(typeof (archiveApi as { resume?: unknown }).resume).toBe("function");
  });

  it("POSTs /api/v1/commercial/archive/{kpId}/resume", async () => {
    mockPost.mockResolvedValue({ draft_id: "draft-resume-42" });

    const resume = (archiveApi as { resume: (kpId: number) => Promise<unknown> }).resume;
    await resume(42);

    expect(mockPost).toHaveBeenCalledWith("/api/v1/commercial/archive/42/resume");
  });
});

describe("archiveApi xlsx_delivery_in_unit", () => {
  it("builds document URL for the embed kind", () => {
    expect(archiveApi.buildDocumentUrl(7, "xlsx_delivery_in_unit")).toBe(
      "/api/v1/commercial/archive/7/files/xlsx_delivery_in_unit",
    );
  });

  it("uses suffix fallback filename for embed xlsx", async () => {
    mockDownload.mockResolvedValue(undefined);
    await archiveApi.downloadDocument(7, "xlsx_delivery_in_unit");
    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/7/files/xlsx_delivery_in_unit",
      "КП_7_с_доставкой_в_цене.xlsx",
    );
  });

  it("keeps ordinary xlsx fallback name", async () => {
    mockDownload.mockResolvedValue(undefined);
    await archiveApi.downloadDocument(7, "xlsx");
    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/7/files/xlsx",
      "КП_7.xlsx",
    );
  });
});
