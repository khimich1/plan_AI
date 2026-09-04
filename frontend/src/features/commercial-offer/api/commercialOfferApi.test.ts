import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import { httpClient } from "@/shared/api/httpClient";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
  },
}));

const mockPost = httpClient.post as unknown as Mock;
const mockPatch = httpClient.patch as unknown as Mock;
const mockDelete = httpClient.delete as unknown as Mock;

const JSON_HEADERS = { "Content-Type": "application/json" };

const draftStub = { draft_id: "draft-append-1" } as CommercialDraftDetails;

beforeEach(() => {
  vi.clearAllMocks();
});

describe("commercialOfferApi.parseSource", () => {
  it("POSTs JSON with product_type and returns lines", async () => {
    const response = {
      product_type: "piles",
      lines: [
        { index: 0, text: "С120.35-12 B25 5", empty: false, ok: true, reason_text: null },
        { index: 1, text: "плохо", empty: false, ok: false, reason_text: "не совпал формат строки" },
      ],
      unparsed_lines: ["плохо"],
    };
    mockPost.mockResolvedValue(response);

    const result = await commercialOfferApi.parseSource({
      text: "С120.35-12 B25 5\nплохо",
      productType: "piles",
    });

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/commercial/parse",
      JSON.stringify({
        text: "С120.35-12 B25 5\nплохо",
        product_type: "piles",
        lint_only: true,
      }),
      JSON_HEADERS,
      undefined,
    );
    expect(result.lines).toHaveLength(2);
    expect(result.lines[1]?.ok).toBe(false);
    expect(result.product_type).toBe("piles");
  });
});

/**
 * MNA-501 / MNA-103 client contract — RED until commercialOfferApi grows append helpers.
 */
describe("commercialOfferApi append / undo / delete (MNA-501)", () => {
  it("startAppendCycle POSTs /append/start with product_type JSON", async () => {
    mockPost.mockResolvedValue(draftStub);

    const result = await commercialOfferApi.startAppendCycle("draft-append-1", "piles");

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/commercial/drafts/draft-append-1/append/start",
      JSON.stringify({ product_type: "piles" }),
      JSON_HEADERS,
    );
    expect(result).toEqual(draftStub);
  });

  it("undoLastAppendBatch POSTs /append/undo-last", async () => {
    mockPost.mockResolvedValue(draftStub);

    const result = await commercialOfferApi.undoLastAppendBatch("draft-append-1");

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/commercial/drafts/draft-append-1/append/undo-last",
    );
    expect(result).toEqual(draftStub);
  });

  it("deleteDraftLine DELETEs /lines/{line_id}", async () => {
    mockDelete.mockResolvedValue(draftStub);

    const result = await commercialOfferApi.deleteDraftLine("draft-append-1", "ln_plates_2");

    expect(mockDelete).toHaveBeenCalledWith(
      "/api/v1/commercial/drafts/draft-append-1/lines/ln_plates_2",
    );
    expect(result).toEqual(draftStub);
  });

  it("patchDraftLine PATCHes /lines/{line_id} with qty JSON", async () => {
    mockPatch.mockResolvedValue(draftStub);
    const result = await commercialOfferApi.patchDraftLine("draft-append-1", "ln_1", { qty: 90 });
    expect(mockPatch).toHaveBeenCalledWith(
      "/api/v1/commercial/drafts/draft-append-1/lines/ln_1",
      JSON.stringify({ qty: 90 }),
      JSON_HEADERS,
    );
    expect(result).toEqual(draftStub);
  });

  it("restoreDraftLines POSTs /lines/restore", async () => {
    mockPost.mockResolvedValue(draftStub);
    const line = { line_id: "ln_1", qty: 2 };
    const result = await commercialOfferApi.restoreDraftLines("draft-append-1", {
      index: 0,
      lines: [line],
      replace_line_ids: ["ln_new"],
    });
    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/commercial/drafts/draft-append-1/lines/restore",
      JSON.stringify({ index: 0, lines: [line], replace_line_ids: ["ln_new"] }),
      JSON_HEADERS,
    );
    expect(result).toEqual(draftStub);
  });
});

describe("commercialOfferApi.ocrPage", () => {
  it("POSTs ocr-page multipart and returns page OCR payload", async () => {
    const payload = {
      normalized_text: "ПБ 34-15-10п 15\nПБ 60-12-8п 3",
      ocr_verify_failed: false,
      ocr_corrections: [{ action: "replaced", reason: "qty" }],
    };
    mockPost.mockResolvedValue(payload);
    const image = new File(["png"], "page.png", { type: "image/png" });

    const result = await commercialOfferApi.ocrPage("draft-ocr-1", image);

    expect(mockPost).toHaveBeenCalledTimes(1);
    const [url, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(url).toBe("/api/v1/commercial/drafts/draft-ocr-1/ocr-page");
    expect(body).toBeInstanceOf(FormData);
    expect(body.get("image")).toBe(image);
    expect(body.has("mode")).toBe(false);
    expect(body.has("text")).toBe(false);
    expect(mockPatch).not.toHaveBeenCalled();
    expect(result).toEqual(payload);
  });
});
