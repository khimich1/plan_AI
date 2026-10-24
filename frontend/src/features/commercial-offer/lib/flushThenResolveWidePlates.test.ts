import { describe, expect, it, vi } from "vitest";

import {
  buildMergedFlushText,
  buildWidePlateApplyFlushText,
  buildWidePlateResolveDecisions,
  flushThenResolveWidePlates,
  normalizeLineKey,
  pruneWidePlateActions,
} from "@/features/commercial-offer/lib/flushThenResolveWidePlates";
import { liveWidePlateLines } from "@/features/commercial-offer/lib/liveWidePlateLines";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

describe("flushThenResolveWidePlates", () => {
  it("normalizeLineKey aligns compact OCR with ПБ…п marks", () => {
    expect(normalizeLineKey("68-15-8 2")).toBe("68-15-8");
    expect(normalizeLineKey("ПБ 68-15-8п 2")).toBe("68-15-8");
    expect(normalizeLineKey("Плиты ПБ 27-15-8п 2")).toBe("27-15-8");
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

  it("maps replace from OCR line without «п» onto flushed ПБ…п line", () => {
    const replacementText = "ПБ 76-12-8 13\nПБ 76-3.0-8 13";
    const decisions = buildWidePlateResolveDecisions({
      liveLines: [{ id: "live-wide-0", line: "ПБ 76-15-8 13", qty: 13 }],
      flushedWideLines: [{ id: "wide-1", line: "ПБ 76-15-8п 13", qty: 13 }],
      decisionsById: {
        "live-wide-0": { action: "replace", replacementText },
      },
    });

    expect(decisions).toEqual([
      {
        lineId: "wide-1",
        sourceLine: "ПБ 76-15-8п 13",
        action: "replace",
        replacementText,
      },
    ]);
  });

  it("matches a line with шт to the canonical mark without шт", () => {
    expect(normalizeLineKey("ПБ 76-15-8п 13 шт")).toBe(normalizeLineKey("ПБ 76-15-8п 13"));
    expect(normalizeLineKey("ПБ 76-15-8 13 шт.")).toBe(normalizeLineKey("76-15-8"));

    const decisions = buildWidePlateResolveDecisions({
      liveLines: [{ id: "live-wide-0", line: "ПБ 76-15-8 13 шт", qty: 13 }],
      flushedWideLines: [{ id: "wide-1", line: "ПБ 76-15-8п 13", qty: 13 }],
      decisionsById: {
        "live-wide-0": {
          action: "replace",
          replacementText: "ПБ 76-12-8 13",
        },
      },
    });

    expect(decisions[0]?.lineId).toBe("wide-1");
    expect(decisions[0]?.action).toBe("replace");
    expect(decisions[0]?.sourceLine).toBe("ПБ 76-15-8п 13");
  });

  it("does not send an unmatched replace as confirm", async () => {
    const unmatched = {
      liveLines: [{ id: "live-wide-0", line: "ПБ 76-15-8 13", qty: 13 }],
      flushedWideLines: [{ id: "wide-9", line: "ПБ 50-15-8п 1", qty: 1 }],
      decisionsById: {
        "live-wide-0": {
          action: "replace" as const,
          replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
        },
      },
    };
    expect(() => buildWidePlateResolveDecisions(unmatched)).toThrow(
      "Не удалось применить решение по широким плитам: позиция не найдена в списке.",
    );

    const resolveWidePlates = vi.fn(async () => undefined);
    await expect(
      flushThenResolveWidePlates({
        draftId: "d1",
        flushText: "ПБ 76-15-8 13",
        persistedText: "ПБ 76-15-8 13",
        liveLines: [{ id: "live-wide-0", line: "ПБ 76-15-8 13", qty: 13 }],
        decisionsById: {
          "live-wide-0": {
            action: "replace",
            replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
          },
        },
        currentWideLines: [{ id: "wide-9", line: "ПБ 50-15-8п 1", qty: 1 }],
        updateInput: vi.fn(async () => {
          throw new Error("updateInput must not run");
        }),
        resolveWidePlates,
      }),
    ).rejects.toThrow("Не удалось применить решение по широким плитам: позиция не найдена в списке.");
    expect(resolveWidePlates).not.toHaveBeenCalled();
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

  const fourLineReplacement = [
    "ПБ 76-12-8 13",
    "ПБ 76-3.0-8 13",
    "ПБ 74-12-8 4",
    "ПБ 74-3.0-8 4",
  ].join("\n");

  it("sends a four-line replacementText in full", () => {
    const decisions = buildWidePlateResolveDecisions({
      liveLines: [{ id: "live-wide-0", line: "ПБ 76-15-8 13", qty: 13 }],
      flushedWideLines: [{ id: "wide-1", line: "ПБ 76-15-8п 13", qty: 13 }],
      decisionsById: {
        "live-wide-0": { action: "replace", replacementText: fourLineReplacement },
      },
    });

    expect(decisions[0]?.replacementText).toBe(fourLineReplacement);
    expect(decisions[0]?.replacementText.split("\n")).toHaveLength(4);
  });

  it("flushes each photo page when the store holds the last batch", () => {
    const first = "ПБ 76-15-8 13\nПБ 60-12-8 1";
    const second = "ПБ 58-12-8 2";
    const flushText = buildWidePlateApplyFlushText({
      hasStarted: true,
      pages: [
        { id: "p1", batchReviewText: first },
        { id: "p2", batchReviewText: second },
      ],
      activePageId: "p1",
      storeText: second,
    });

    expect(flushText).toBe(`${first}\n${second}`);
  });

  it("does not replace-update a photo queue when the active page was edited", async () => {
    const updateInput = vi.fn(async () => {
      throw new Error("updateInput must not run");
    });
    const resolveWidePlates = vi.fn(async () => undefined);

    await flushThenResolveWidePlates({
      draftId: "d1",
      flushText: "ПБ 76-15-8 13\nПБ 60-12-8 5\nПБ 58-12-8 2",
      persistedText: "ПБ 76-15-8 13\nПБ 60-12-8 1\nПБ 58-12-8 2",
      photoQueue: true,
      liveLines: [{ id: "live-wide-0", line: "ПБ 76-15-8 13", qty: 13 }],
      decisionsById: {
        "live-wide-0": { action: "replace", replacementText: fourLineReplacement },
      },
      currentWideLines: [{ id: "wide-1", line: "ПБ 76-15-8п 13", qty: 13 }],
      updateInput,
      resolveWidePlates,
    });

    expect(updateInput).not.toHaveBeenCalled();
    expect(resolveWidePlates).toHaveBeenCalledWith({
      draftId: "d1",
      decisions: [
        expect.objectContaining({
          lineId: "wide-1",
          action: "replace",
          replacementText: fourLineReplacement,
        }),
      ],
    });
  });

  it("does not replace-update one photo when its text matches input_text", async () => {
    const text = "ПБ 76-15-8 13";
    const flushText = buildWidePlateApplyFlushText({
      hasStarted: true,
      pages: [{ id: "p1", batchReviewText: text }],
      activePageId: "p1",
      storeText: text,
    });
    expect(flushText).toBe(text);

    const updateInput = vi.fn();
    const resolveWidePlates = vi.fn(async () => undefined);
    await flushThenResolveWidePlates({
      draftId: "d1",
      flushText,
      persistedText: text,
      photoQueue: true,
      liveLines: [{ id: "live-wide-0", line: text, qty: 13 }],
      decisionsById: {
        "wide-1": {
          action: "replace",
          replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
        },
      },
      currentWideLines: [{ id: "wide-1", line: "ПБ 76-15-8п 13", qty: 13 }],
      updateInput,
      resolveWidePlates,
    });

    expect(updateInput).not.toHaveBeenCalled();
    expect(resolveWidePlates).toHaveBeenCalled();
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

  it("resends replace when server wide lines are empty but the page still has the mark", async () => {
    const pageLine = "ПБ 76-15-8 13";
    const updateInput = vi.fn(async () => {
      throw new Error("updateInput must not run");
    });
    const resolveWidePlates = vi.fn(async () => undefined);

    await flushThenResolveWidePlates({
      draftId: "d1",
      flushText: `${pageLine}\nПБ 60-12-8 1`,
      persistedText: `${pageLine}\nПБ 60-12-8 1`,
      photoQueue: true,
      liveLines: [{ id: "live-wide-0", line: pageLine, qty: 13 }],
      decisionsById: {
        [normalizeLineKey(pageLine)]: { action: "replace", replacementText: fourLineReplacement },
      },
      currentWideLines: [],
      updateInput,
      resolveWidePlates,
    });

    expect(updateInput).not.toHaveBeenCalled();
    expect(resolveWidePlates).toHaveBeenCalledTimes(1);
    expect(resolveWidePlates).toHaveBeenCalledWith({
      draftId: "d1",
      decisions: [
        expect.objectContaining({
          action: "replace",
          sourceLine: pageLine,
          replacementText: fourLineReplacement,
        }),
      ],
    });
  });

  it("drops replace when the page text no longer contains that mark", () => {
    const kept = pruneWidePlateActions(
      {
        "wide-0": {
          action: "replace",
          replacementText: fourLineReplacement,
          sourceLine: "ПБ 76-15-8 13",
        },
      },
      "ПБ 60-12-8 2",
    );

    expect(kept).toEqual({});
  });

  it("keeps replace while the page still shows the same mark", () => {
    const decision = {
      action: "replace" as const,
      replacementText: fourLineReplacement,
      sourceLine: "ПБ 76-15-8п 13",
    };

    expect(pruneWidePlateActions({ "wide-0": decision }, "ПБ 76-15-8 13")).toEqual({
      "wide-0": decision,
    });
  });

  it("does not clear choices when the page text is empty", () => {
    const actions = {
      "wide-0": {
        action: "replace" as const,
        replacementText: fourLineReplacement,
        sourceLine: "ПБ 76-15-8 13",
      },
    };

    expect(pruneWidePlateActions(actions, " \n")).toBe(actions);
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
