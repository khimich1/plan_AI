import { describe, expect, it, vi } from "vitest";

import {
  buildMergedFlushText,
  buildWidePlateResolveDecisions,
  flushThenResolveWidePlates,
  normalizeLineKey,
} from "@/features/commercial-offer/lib/flushThenResolveWidePlates";
import { liveWidePlateLines } from "@/features/commercial-offer/lib/liveWidePlateLines";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

describe("flushThenResolveWidePlates", () => {
  it("normalizeLineKey aligns compact OCR with ПБ…п marks", () => {
    expect(normalizeLineKey("68-15-8 2")).toBe("68-15-8 2");
    expect(normalizeLineKey("ПБ 68-15-8п 2")).toBe("68-15-8 2");
    expect(normalizeLineKey("Плиты ПБ 27-15-8п 2")).toBe("27-15-8 2");
  });
  it("builds resolve payload from live lines after flush text", () => {
    const editorText = "34-15-10п 15";
    const liveLines = liveWidePlateLines(editorText);
    const decisions = buildWidePlateResolveDecisions({
      liveLines,
      flushedWideLines: [{ id: "wide-after-flush", line: "34-15-10п 15", qty: 15 }],
      decisionsById: {
        "live-wide-0": {
          action: "replace",
          replacementText: "ПБ 34-12-10п 15\nПБ 34-3.0-10п 15",
        },
      },
    });

    expect(decisions).toHaveLength(1);
    expect(decisions[0]?.sourceLine).toBe("34-15-10п 15");
    expect(decisions[0]?.sourceLine).not.toBe("44-15-10п 5");
    expect(decisions[0]?.lineId).toBe("wide-after-flush");
    expect(decisions[0]?.action).toBe("replace");
    expect(decisions[0]?.replacementText).toContain("15");
  });

  it("maps exclude from compact live line without «п» onto flushed ПБ…п line", () => {
    const decisions = buildWidePlateResolveDecisions({
      liveLines: liveWidePlateLines("68-15-8 2\n27-15-8 2"),
      flushedWideLines: [
        { id: "wide-1", line: "ПБ 68-15-8п 2", qty: 2 },
        { id: "wide-2", line: "ПБ 27-15-8п 2", qty: 2 },
      ],
      decisionsById: {
        "live-wide-0": { action: "exclude", replacementText: "" },
        "live-wide-1": { action: "exclude", replacementText: "" },
      },
    });

    expect(decisions).toEqual([
      expect.objectContaining({ lineId: "wide-1", action: "exclude", sourceLine: "ПБ 68-15-8п 2" }),
      expect.objectContaining({ lineId: "wide-2", action: "exclude", sourceLine: "ПБ 27-15-8п 2" }),
    ]);
  });

  it("multi-page merge does not drop other pages", () => {
    const merged = buildMergedFlushText({
      hasStarted: true,
      pages: [
        { id: "p1", batchReviewText: "ПБ 60-12-8п 2" },
        { id: "p2", batchReviewText: "44-15-10п 5" },
      ],
      activePageId: "p2",
      editorText: "34-15-10п 15",
      singlePageText: "34-15-10п 15",
    });

    expect(merged).toContain("ПБ 60-12-8п 2");
    expect(merged).toContain("34-15-10п 15");
    expect(merged).not.toContain("44-15-10п 5");
  });

  it("keeps confirm on 15 dm live line", () => {
    const decisions = buildWidePlateResolveDecisions({
      liveLines: liveWidePlateLines("34-15-10п 15"),
      flushedWideLines: [{ id: "wide-15", line: "34-15-10п 15", qty: 15 }],
      decisionsById: {},
    });

    expect(decisions[0]?.action).toBe("confirm");
    expect(decisions[0]?.sourceLine).toBe("34-15-10п 15");
  });

  it("calls updateInput before resolveWidePlates", async () => {
    const order: string[] = [];
    const updateInput = vi.fn(async () => {
      order.push("updateInput");
      return {
        draft_id: "d1",
        metadata: {
          input_text: "34-15-10п 15",
          normalized_text: "34-15-10п 15",
          wide_plate_lines: [{ id: "wide-after-flush", line: "34-15-10п 15", qty: 15 }],
        },
      } as CommercialDraftDetails;
    });
    const resolveWidePlates = vi.fn(async () => {
      order.push("resolveWidePlates");
    });

    await flushThenResolveWidePlates({
      draftId: "d1",
      flushText: "34-15-10п 15",
      persistedText: "44-15-10п 5",
      liveLines: liveWidePlateLines("34-15-10п 15"),
      decisionsById: {
        "live-wide-0": { action: "confirm", replacementText: "" },
      },
      updateInput,
      resolveWidePlates,
    });

    expect(order).toEqual(["updateInput", "resolveWidePlates"]);
    expect(updateInput).toHaveBeenCalledWith({
      draftId: "d1",
      text: "34-15-10п 15",
      image: null,
      mode: "replace",
    });
    expect(resolveWidePlates).toHaveBeenCalledWith({
      draftId: "d1",
      decisions: [
        expect.objectContaining({
          lineId: "wide-after-flush",
          sourceLine: "34-15-10п 15",
          action: "confirm",
        }),
      ],
    });
  });

  it("does not resolve if flush fails", async () => {
    const resolveWidePlates = vi.fn();
    await expect(
      flushThenResolveWidePlates({
        draftId: "d1",
        flushText: "34-15-10п 15",
        persistedText: "44-15-10п 5",
        liveLines: liveWidePlateLines("34-15-10п 15"),
        decisionsById: {},
        updateInput: async () => {
          throw new Error("flush failed");
        },
        resolveWidePlates,
      }),
    ).rejects.toThrow("flush failed");
    expect(resolveWidePlates).not.toHaveBeenCalled();
  });
});
