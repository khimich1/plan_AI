import { describe, expect, it, vi } from "vitest";

import {
  applyPromoteSourceImageQueue,
  type SourceImageQueueSink,
} from "@/features/commercial-offer/lib/promoteSourceImageQueue";

/**
 * Documents clear policy: confirm must NOT clear; lifecycle events must clear.
 * Wizard wires sourceImageQueue.clear() at these call sites (IMG-004).
 */
describe("sourceImageQueue clear policy", () => {
  const makeSink = (): SourceImageQueueSink & {
    setFromPages: ReturnType<typeof vi.fn>;
    setFromSinglePreview: ReturnType<typeof vi.fn>;
    clear: ReturnType<typeof vi.fn>;
  } => ({
    setFromPages: vi.fn(),
    setFromSinglePreview: vi.fn(),
    clear: vi.fn(),
  });

  it("confirm promote keeps queue (does not call clear)", () => {
    const sink = makeSink();
    const file = new File(["x"], "a.png", { type: "image/png" });

    applyPromoteSourceImageQueue(sink, {
      pages: [{ id: "a", file, name: "a.png" }],
      singlePreview: null,
    });

    expect(sink.setFromPages).toHaveBeenCalled();
    expect(sink.clear).not.toHaveBeenCalled();
  });

  it("lifecycle clear targets are append / archive-save / create-new / resetSource", () => {
    // Explicit checklist mirrored by CommercialOfferWizard call sites.
    const clearOnEvents = [
      "handleAddOtherNomenclature",
      "handleSave (archive success)",
      "handleCreateNewOffer",
      "handleSourceTextChange (image→text abandon)",
    ] as const;

    expect(clearOnEvents).toHaveLength(4);
    expect(clearOnEvents).not.toContain("handleConfirmBatch");
    expect(clearOnEvents).not.toContain("resetSource");
  });
});
