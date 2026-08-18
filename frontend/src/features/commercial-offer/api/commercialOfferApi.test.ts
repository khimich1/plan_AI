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
const mockDelete = httpClient.delete as unknown as Mock;

const JSON_HEADERS = { "Content-Type": "application/json" };

const draftStub = { draft_id: "draft-append-1" } as CommercialDraftDetails;

beforeEach(() => {
  vi.clearAllMocks();
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
});
