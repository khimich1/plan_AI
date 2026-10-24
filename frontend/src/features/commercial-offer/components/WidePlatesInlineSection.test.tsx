import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { WidePlatesInlineSection } from "@/features/commercial-offer/components/WidePlatesInlineSection";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";

const fourLineReplacement = [
  "ПБ 76-12-8 13",
  "ПБ 76-3.0-8 13",
  "ПБ 74-12-8 4",
  "ПБ 74-3.0-8 4",
].join("\n");

const draft = {
  metadata: {
    wide_plates_resolved: false,
    wide_plate_lines: [
      { id: "live-wide-0", line: "ПБ 76-15-8 13", qty: 13 },
      { id: "live-wide-1", line: "ПБ 65-14-8 2", qty: 2 },
    ],
  },
} as CommercialDraftDetails;

describe("WidePlatesInlineSection choice by mark", () => {
  afterEach(() => {
    cleanup();
  });

  it("shows replace stored under wide-0 on live-wide-0 with the same mark", () => {
    render(
      <WidePlatesInlineSection
        draft={draft}
        decisions={{
          "wide-0": {
            action: "replace",
            replacementText: fourLineReplacement,
            sourceLine: "ПБ 76-15-8п 13",
          },
        }}
        isPending={false}
        onDecisionChange={vi.fn()}
        onApply={vi.fn()}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: /позиций требуют внимания/i }));

    const replaceRadios = screen.getAllByRole("radio", { name: "Заменить" });
    const confirmRadios = screen.getAllByRole("radio", { name: "Подтвердить" });
    expect(replaceRadios[0]).toBeChecked();
    expect(replaceRadios[1]).not.toBeChecked();
    expect(confirmRadios[1]).toBeChecked();
    expect(screen.getByRole("textbox", { name: /Заменяющие позиции/ })).toHaveValue(fourLineReplacement);
    expect(screen.getByRole("textbox", { name: /Заменяющие позиции/ })).not.toHaveValue("");
  });
});
