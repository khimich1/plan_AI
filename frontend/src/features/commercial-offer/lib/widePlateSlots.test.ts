import { describe, expect, it } from "vitest";

import { buildAutoSplitSuggestion } from "@/features/commercial-offer/lib/widePlateSuggestion";
import {
  applyWideSlot,
  carryWideSlots,
  chooseWideSlotAction,
  commitWideSlots,
  photoQueueWidePlateBlocksFinish,
  prepareTextReviewConfirm,
  routeWidePlateApply,
  TEXT_REVIEW_SLOT_PAGE_ID,
  uniqueWideSlotsByLine,
  type TextReviewConfirmPlan,
} from "@/features/commercial-offer/lib/widePlateSlots";

describe("carryWideSlots", () => {
  it("keeps the slot when the only change is the text of that line", () => {
    const previousLines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13"];
    const nextLines = ["ПБ 10-12-8 1", "ПБ 68-15-8 13"];
    const slot = {
      id: "slot-wide-1",
      lineIndex: 1,
      action: "replace" as const,
      replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
      accepted: false,
    };

    const next = carryWideSlots(previousLines, nextLines, [slot]);

    expect(next).toEqual([
      {
        id: "slot-wide-1",
        lineIndex: 1,
        action: "replace",
        replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
        accepted: false,
      },
    ]);
    expect(next[0]?.id).not.toMatch(/^live-wide-/);
  });

  it("keeps the slot on the same plate when a line is inserted above", () => {
    const previousLines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13"];
    const nextLines = ["ПБ 8-12-8 1", "ПБ 10-12-8 1", "ПБ 76-15-8 13"];
    const slot = {
      id: "slot-wide-b",
      lineIndex: 1,
      action: "replace" as const,
      replacementText: "ПБ 76-12-8 13",
      accepted: false,
    };

    const next = carryWideSlots(previousLines, nextLines, [slot]);

    expect(next).toEqual([{ ...slot, lineIndex: 2 }]);
  });

  it("edits the first of two identical marks and leaves the second slot", () => {
    const previousLines = ["ПБ 76-15-8 13", "ПБ 76-15-8 13"];
    const nextLines = ["ПБ 68-15-8 13", "ПБ 76-15-8 13"];
    const first = {
      id: "slot-first",
      lineIndex: 0,
      action: "replace" as const,
      replacementText: "первая",
      accepted: false,
    };
    const second = {
      id: "slot-second",
      lineIndex: 1,
      action: "exclude" as const,
      replacementText: "вторая",
      accepted: false,
    };

    const next = carryWideSlots(previousLines, nextLines, [first, second]);

    expect(next).toEqual([
      { ...first, lineIndex: 0 },
      { ...second, lineIndex: 1 },
    ]);
  });

  it("drops the slot when the line is no longer wider than 12 dm", () => {
    const previousLines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13"];
    const nextLines = ["ПБ 10-12-8 1", "ПБ 76-12-8 13"];
    const replacementText = "ПБ 76-12-8 13\nПБ 76-3.0-8 13";

    const next = carryWideSlots(previousLines, nextLines, [
      {
        id: "slot-wide-1",
        lineIndex: 1,
        action: "replace",
        replacementText,
        accepted: false,
      },
    ]);

    expect(next).toEqual([]);
    expect(nextLines).toEqual(["ПБ 10-12-8 1", "ПБ 76-12-8 13"]);
    expect(nextLines.join("\n")).not.toContain(replacementText);
  });

  it("drops the slot when its line is deleted", () => {
    const previousLines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13", "ПБ 20-12-8 2"];
    const nextLines = ["ПБ 10-12-8 1", "ПБ 20-12-8 2"];

    const next = carryWideSlots(previousLines, nextLines, [
      {
        id: "slot-wide-1",
        lineIndex: 1,
        action: "replace",
        replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
        accepted: false,
      },
    ]);

    expect(next).toEqual([]);
    expect(nextLines).toEqual(["ПБ 10-12-8 1", "ПБ 20-12-8 2"]);
  });

  it("clears accepted when a confirmed line is edited into another wide mark", () => {
    const previousLines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13"];
    const nextLines = ["ПБ 10-12-8 1", "ПБ 68-15-8 13"];
    const slot = {
      id: "slot-wide-1",
      lineIndex: 1,
      action: "confirm" as const,
      replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
      accepted: true,
    };

    const next = carryWideSlots(previousLines, nextLines, [slot]);

    expect(next).toEqual([
      {
        id: "slot-wide-1",
        lineIndex: 1,
        action: "confirm",
        replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
        accepted: false,
      },
    ]);
  });

  it("keeps accepted when the confirmed line text is unchanged", () => {
    const lines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13"];
    const slot = {
      id: "slot-wide-1",
      lineIndex: 1,
      action: "confirm" as const,
      replacementText: "ПБ 76-12-8 13",
      accepted: true,
    };

    expect(carryWideSlots(lines, [...lines], [slot])).toEqual([slot]);
  });
});

describe("applyWideSlot", () => {
  it("replaces that line with the non-empty field lines", () => {
    const lines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13", "ПБ 20-12-8 2"];

    const result = applyWideSlot(lines, 1, {
      action: "replace",
      replacementText: "ПБ 76-12-8 13\n\n  \nПБ 76-3.0-8 13",
      accepted: false,
    });

    expect(result).toEqual({
      ok: true,
      lines: ["ПБ 10-12-8 1", "ПБ 76-12-8 13", "ПБ 76-3.0-8 13", "ПБ 20-12-8 2"],
    });
    expect(result.ok && result.lines).not.toContain("ПБ 76-15-8 13");
  });

  it("removes only the excluded line", () => {
    const lines = ["ПБ 76-15-8 13", "ПБ 76-15-8 13", "ПБ 20-12-8 2"];

    const result = applyWideSlot(lines, 1, {
      action: "exclude",
      replacementText: "ПБ 76-12-8 13",
      accepted: false,
    });

    expect(result).toEqual({
      ok: true,
      lines: ["ПБ 76-15-8 13", "ПБ 20-12-8 2"],
    });
  });

  it("confirm returns the same lines", () => {
    const lines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13", "ПБ 20-12-8 2"];

    const result = applyWideSlot(lines, 1, {
      action: "confirm",
      replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
      accepted: true,
    });

    expect(result).toEqual({ ok: true, lines: [...lines] });
  });

  it("rejects an empty replacement and leaves the lines", () => {
    const lines = ["ПБ 10-12-8 1", "ПБ 76-15-8 13"];
    const decision = { action: "replace" as const, accepted: false };

    expect(applyWideSlot(lines, 1, { ...decision, replacementText: "" })).toEqual({
      ok: false,
      error: "Укажите заменяющие позиции.",
    });
    expect(applyWideSlot(lines, 1, { ...decision, replacementText: " \n\n  " })).toEqual({
      ok: false,
      error: "Укажите заменяющие позиции.",
    });
    expect(lines).toEqual(["ПБ 10-12-8 1", "ПБ 76-15-8 13"]);
  });
});

describe("commitWideSlots", () => {
  it("commits slots in order and stops at the first empty replace", () => {
    const lines = ["ПБ 76-15-8 13", "ПБ 68-15-8 13", "ПБ 65-14-8 2"];
    const result = commitWideSlots(lines, [
      {
        id: "slot-exclude",
        lineIndex: 0,
        action: "exclude",
        replacementText: "",
        accepted: false,
      },
      {
        id: "slot-replace",
        lineIndex: 1,
        action: "replace",
        replacementText: " \n",
        accepted: false,
      },
      {
        id: "slot-confirm",
        lineIndex: 2,
        action: "confirm",
        replacementText: "",
        accepted: false,
      },
    ]);

    expect(result.lines).toEqual(["ПБ 68-15-8 13", "ПБ 65-14-8 2"]);
    expect(result.failedId).toBe("slot-replace");
    expect(result.error).toBe("Укажите заменяющие позиции.");
    expect(result.slots.find((slot) => slot.id === "slot-confirm")?.accepted).toBe(false);
    expect(result.slots.find((slot) => slot.id === "slot-replace")?.accepted).toBe(false);
    expect(result.lines[result.slots.find((slot) => slot.id === "slot-replace")?.lineIndex ?? -1]).toBe(
      "ПБ 68-15-8 13",
    );
  });

  it("applies one slot and keeps the other wide line on its original slot", () => {
    const lines = ["ПБ 68-15-8 15", "ПБ 76-15-8 13"];
    const result = commitWideSlots(
      lines,
      [
        {
          id: "slot-68",
          lineIndex: 0,
          action: "replace",
          replacementText: "ПБ 68-12-8 15\nПБ 68-3.0-8 15",
          accepted: false,
        },
        {
          id: "slot-76",
          lineIndex: 1,
          action: "replace",
          replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
          accepted: false,
        },
      ],
      ["slot-68"],
    );

    expect(result.lines).toEqual(["ПБ 68-12-8 15", "ПБ 68-3.0-8 15", "ПБ 76-15-8 13"]);
    expect(result.error).toBeNull();
    expect(result.slots.map((slot) => slot.lineIndex)).toEqual([2]);
    expect(result.slots.map((slot) => slot.id)).toEqual(["slot-76"]);
    expect(result.slots[0]?.accepted).toBe(false);
  });
});

describe("uniqueWideSlotsByLine", () => {
  it("keeps the filled slot when a fresh confirm copy shares the line", () => {
    const filled = {
      id: "slot-filled",
      lineIndex: 0,
      action: "replace" as const,
      replacementText: "ПБ 68-12-8 15",
      accepted: false,
    };
    const fresh = {
      id: "slot-fresh",
      lineIndex: 0,
      action: "confirm" as const,
      replacementText: "",
      accepted: false,
    };

    expect(uniqueWideSlotsByLine([fresh, filled]).map((slot) => slot.id)).toEqual(["slot-filled"]);
  });
});

describe("chooseWideSlotAction", () => {
  it("fills the suggestion only when replace is chosen and the field is empty", () => {
    const line = "ПБ 76-15-8 13";
    const slot = {
      id: "slot-1",
      lineIndex: 0,
      action: "confirm" as const,
      replacementText: "",
      accepted: false,
    };

    const next = chooseWideSlotAction(slot, "replace", line, 13);

    expect(next).toEqual({
      ...slot,
      action: "replace",
      replacementText: buildAutoSplitSuggestion(line, 13),
    });
  });

  it("does not rewrite a field that already has text when replace is chosen again", () => {
    const slot = {
      id: "slot-1",
      lineIndex: 0,
      action: "confirm" as const,
      replacementText: "ручной текст",
      accepted: false,
    };

    const next = chooseWideSlotAction(slot, "replace", "ПБ 68-15-8 13", 13);

    expect(next.replacementText).toBe("ручной текст");
    expect(next.action).toBe("replace");
  });

  it("carry does not refresh the replacement field", () => {
    const next = carryWideSlots(["ПБ 76-15-8 13"], ["ПБ 68-15-8 13"], [
      {
        id: "slot-1",
        lineIndex: 0,
        action: "replace",
        replacementText: "старое поле",
        accepted: false,
      },
    ]);

    expect(next[0]?.replacementText).toBe("старое поле");
    expect(next[0]?.replacementText).not.toBe(buildAutoSplitSuggestion("ПБ 68-15-8 13", 13));
  });
});

describe("photoQueueWidePlateBlocksFinish", () => {
  const wideText = "ПБ 76-15-8 13";
  const openSlot = {
    id: "slot-1",
    lineIndex: 0,
    action: "confirm" as const,
    replacementText: "",
    accepted: false,
  };
  const acceptedSlot = { ...openSlot, accepted: true };

  it("blocks when a page has an open wide slot", () => {
    expect(photoQueueWidePlateBlocksFinish([{ text: wideText, slots: [openSlot] }])).toBe(true);
  });

  it("does not block when that wide slot is accepted", () => {
    expect(photoQueueWidePlateBlocksFinish([{ text: wideText, slots: [acceptedSlot] }])).toBe(false);
  });

  it("blocks when another page still has an open slot", () => {
    expect(
      photoQueueWidePlateBlocksFinish([
        { text: wideText, slots: [acceptedSlot] },
        { text: wideText, slots: [{ ...openSlot, id: "slot-2" }] },
      ]),
    ).toBe(true);
  });

  it("blocks a wide line that has no slot yet", () => {
    expect(photoQueueWidePlateBlocksFinish([{ text: wideText, slots: [] }])).toBe(true);
  });

  it("does not block text without a wide line", () => {
    expect(photoQueueWidePlateBlocksFinish([{ text: "ПБ 76-12-8 13", slots: [] }])).toBe(false);
  });
});

describe("routeWidePlateApply", () => {
  it("photo queue applies locally", () => {
    expect(routeWidePlateApply(true)).toBe("local");
    expect(routeWidePlateApply(false)).toBe("resolve");
  });

  it("open text review applies locally and a confirmed photo queue still resolves", () => {
    expect(routeWidePlateApply(false, { textReview: true })).toBe("local");
    expect(routeWidePlateApply(false)).toBe("resolve");
    expect(TEXT_REVIEW_SLOT_PAGE_ID).toBe("text-review");
  });
});

describe("prepareTextReviewConfirm", () => {
  const closedPlan = (plan: TextReviewConfirmPlan) => {
    if (!plan.closeReview) {
      throw new Error(plan.error);
    }
    return plan;
  };
  const openSlot = (
    id: string,
    lineIndex: number,
    action: "confirm" | "replace" | "exclude",
    replacementText = "",
  ) => ({
    id,
    lineIndex,
    action,
    replacementText,
    accepted: false,
  });

  it("writes open replacements and keeps the typed neighbor", () => {
    const editor = "ПБ 76-15-8 13\nПБ 76-12-8п 18";
    const plan = prepareTextReviewConfirm(
      editor,
      [openSlot("slot-76", 0, "replace", "ПБ 76-12-8п 13\nПБ 76-3.0-8п 13")],
      editor,
    );

    const saved = closedPlan(plan);
    expect(saved.uploadText).toBe("ПБ 76-12-8п 13\nПБ 76-3.0-8п 13\nПБ 76-12-8п 18");
    expect(saved.uploadText).not.toContain("76-15-8");
    expect(saved.text).toContain("ПБ 76-12-8п 18");
  });

  it("drops an excluded line before comparing with the server batch", () => {
    const onlyWide = "ПБ 68-15-8 15";
    const cleared = prepareTextReviewConfirm(onlyWide, [openSlot("slot-only", 0, "exclude")], onlyWide);
    const clearedPlan = closedPlan(cleared);
    expect(clearedPlan.text.trim()).toBe("");
    expect(clearedPlan.uploadText).toBeNull();

    const editor = "ПБ 68-15-8 15\nПБ 68-12-8п 3";
    const kept = closedPlan(prepareTextReviewConfirm(editor, [openSlot("slot-68", 0, "exclude")], editor));
    expect(kept.uploadText).toBe("ПБ 68-12-8п 3");
    expect(kept.uploadText).not.toContain("68-15-8");
  });

  it("keeps confirmed lines and skips upload when the batch already matches", () => {
    const editor = "ПБ 76-15-8 13";
    const plan = prepareTextReviewConfirm(editor, [openSlot("slot-keep", 0, "confirm")], editor);

    const saved = closedPlan(plan);
    expect(saved.uploadText).toBeNull();
    expect(saved.text).toBe(editor);
  });

  it("does not apply an accepted replacement again", () => {
    const editor = "ПБ 76-12-8п 13";
    const plan = prepareTextReviewConfirm(
      editor,
      [
        {
          id: "slot-done",
          lineIndex: 0,
          action: "replace",
          replacementText: "ПБ 76-12-8п 13\nПБ 76-3.0-8п 13",
          accepted: true,
        },
      ],
      editor,
    );

    const saved = closedPlan(plan);
    expect(saved.text).toBe(editor);
    expect(saved.uploadText).toBeNull();
    expect(saved.text).not.toContain("3.0");
  });

  it("stops on an empty replacement and leaves the review open", () => {
    const editor = "ПБ 76-15-8 13\nПБ 68-15-8 15";
    const plan = prepareTextReviewConfirm(
      editor,
      [
        openSlot("slot-76", 0, "replace", "ПБ 76-12-8п 13"),
        openSlot("slot-68", 1, "replace", "  "),
      ],
      editor,
    );

    expect(plan.closeReview).toBe(false);
    if (plan.closeReview) {
      return;
    }
    expect(plan.error).toBe("Укажите заменяющие позиции.");
    expect(plan.text).toBe("ПБ 76-12-8п 13\nПБ 68-15-8 15");
    expect(plan.text).not.toContain("68-12");
    expect(plan.slots.find((slot) => slot.id === "slot-68")?.accepted).toBe(false);
  });
});
