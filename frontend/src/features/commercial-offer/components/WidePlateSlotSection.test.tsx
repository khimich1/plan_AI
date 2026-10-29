import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { WidePlateSlotSection } from "@/features/commercial-offer/components/WidePlateSlotSection";
import type { WideSlot } from "@/features/commercial-offer/lib/widePlateSlots";

const lines = ["ПБ 76-15-8 13", "ПБ 65-14-8 2"];

const openSlot: WideSlot = {
  id: "slot-open",
  lineIndex: 0,
  action: "replace",
  replacementText: "ПБ 76-12-8 13\nПБ 76-3.0-8 13",
  accepted: false,
};

const acceptedSlot: WideSlot = {
  id: "slot-accepted",
  lineIndex: 1,
  action: "confirm",
  replacementText: "",
  accepted: true,
};

const expandAttention = () => {
  fireEvent.click(screen.getByRole("button", { name: /^▸ \d+ позиций требуют внимания$/ }));
};

describe("WidePlateSlotSection", () => {
  afterEach(() => {
    cleanup();
  });

  it("starts collapsed and shows the attention count", () => {
    const { rerender } = render(<WidePlateSlotSection lines={lines} slots={[openSlot]} />);

    expect(screen.getByRole("button", { name: "▸ 1 позиций требуют внимания" })).toBeInTheDocument();
    expect(
      screen.getByText(
        "Разверните блок, чтобы принять решение по каждой позиции, затем нажмите «Применить решения».",
      ),
    ).toBeInTheDocument();
    expect(screen.queryByRole("radio")).not.toBeInTheDocument();
    expect(screen.queryByRole("textbox", { name: /Заменяющие позиции/ })).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Применить решения" })).not.toBeInTheDocument();

    rerender(
      <WidePlateSlotSection
        lines={["ПБ 76-15-8 13", "ПБ 65-14-8 2", "ПБ 80-15-8 1"]}
        slots={[
          openSlot,
          { ...openSlot, id: "slot-second", lineIndex: 1, action: "confirm", replacementText: "" },
          { ...acceptedSlot, lineIndex: 2 },
        ]}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "▸ 2 позиций требуют внимания" }));

    expect(screen.getByRole("button", { name: "▾ 2 позиций требуют внимания" })).toBeInTheDocument();
    expect(screen.getAllByRole("radio").length).toBeGreaterThan(0);
  });

  it("places OK in the radio row", () => {
    render(<WidePlateSlotSection lines={lines} slots={[openSlot]} />);
    expandAttention();

    const exclude = screen.getByRole("radio", { name: "Исключить" });
    const ok = screen.getByRole("button", { name: "ОК" });
    const radioRow = exclude.parentElement?.parentElement;

    expect(radioRow).toContainElement(ok);
    expect(radioRow?.lastElementChild).toBe(ok);
    expect(radioRow).toHaveStyle({ display: "flex", alignItems: "center" });
    expect(screen.getByRole("textbox", { name: /Заменяющие позиции/ }).parentElement).not.toContainElement(ok);
    expect(ok).toHaveAttribute("type", "button");
    expect(ok).not.toHaveStyle({ width: "100%" });
    expect(ok).toHaveStyle({ backgroundColor: "#eef2ff" });
  });

  it("collapses again when collapseKey changes", () => {
    const slots = [openSlot];
    const { rerender } = render(
      <WidePlateSlotSection lines={lines} slots={slots} collapseKey="page-a" />,
    );
    expandAttention();
    expect(screen.getByRole("radio", { name: "Заменить" })).toBeInTheDocument();

    rerender(<WidePlateSlotSection lines={lines} slots={slots} collapseKey="page-b" />);
    expect(screen.queryByRole("radio", { name: "Заменить" })).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "▸ 1 позиций требуют внимания" })).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "▸ 1 позиций требуют внимания" }));
    expect(screen.getByRole("radio", { name: "Заменить" })).toBeInTheDocument();

    rerender(<WidePlateSlotSection lines={lines} slots={slots} collapseKey="page-b" />);
    expect(screen.getByRole("radio", { name: "Заменить" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "▾ 1 позиций требуют внимания" })).toBeInTheDocument();
  });

  it("shows one block per open slot and hides an accepted slot", () => {
    render(<WidePlateSlotSection lines={lines} slots={[openSlot, acceptedSlot]} />);
    expandAttention();

    expect(screen.getByText("ПБ 76-15-8 13")).toBeInTheDocument();
    expect(screen.queryByText("ПБ 65-14-8 2")).not.toBeInTheDocument();
    expect(screen.getByRole("radio", { name: "Заменить" })).toBeChecked();
    expect(screen.getByRole("textbox", { name: /Заменяющие позиции/ })).toHaveValue(openSlot.replacementText);
    expect(screen.getAllByRole("button", { name: "ОК" })).toHaveLength(1);
  });

  it("hides the section when every slot is accepted", () => {
    render(
      <WidePlateSlotSection
        lines={lines}
        slots={[
          { ...openSlot, accepted: true },
          acceptedSlot,
        ]}
      />,
    );

    expect(screen.queryByText("ПБ 76-15-8 13")).not.toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "ОК" })).not.toBeInTheDocument();
  });

  it("hides the replacement field for confirm and exclude", () => {
    render(
      <WidePlateSlotSection
        lines={["ПБ 76-15-8 13", "ПБ 65-14-8 2"]}
        slots={[
          { ...openSlot, id: "slot-confirm", action: "confirm", replacementText: "не показывать" },
          {
            ...openSlot,
            id: "slot-exclude",
            lineIndex: 1,
            action: "exclude",
            replacementText: "тоже не показывать",
          },
        ]}
      />,
    );
    expandAttention();

    expect(screen.queryByRole("textbox", { name: /Заменяющие позиции/ })).not.toBeInTheDocument();
    expect(screen.getAllByRole("button", { name: "ОК" })).toHaveLength(2);
  });

  it("OK reports that slot id", () => {
    const onAccept = vi.fn();
    render(
      <WidePlateSlotSection
        lines={lines}
        slots={[openSlot, { ...acceptedSlot, id: "slot-second", accepted: false }]}
        onAccept={onAccept}
      />,
    );
    expandAttention();

    fireEvent.click(screen.getAllByRole("button", { name: "ОК" })[0]!);

    expect(onAccept).toHaveBeenCalledTimes(1);
    expect(onAccept).toHaveBeenCalledWith("slot-open");
  });

  it("bottom button accepts every open slot id in order", () => {
    const onApply = vi.fn();
    render(
      <WidePlateSlotSection
        lines={lines}
        slots={[
          openSlot,
          { ...openSlot, id: "slot-second", lineIndex: 1, action: "confirm", replacementText: "" },
          acceptedSlot,
        ]}
        onApply={onApply}
      />,
    );
    expandAttention();

    fireEvent.click(screen.getByRole("button", { name: "Применить решения" }));

    expect(onApply).toHaveBeenCalledTimes(1);
    expect(onApply).toHaveBeenCalledWith(["slot-open", "slot-second"]);
  });

  it("radio change reports the slot id and the new action", () => {
    const onAction = vi.fn();
    render(<WidePlateSlotSection lines={lines} slots={[openSlot]} onAction={onAction} />);
    expandAttention();

    fireEvent.click(screen.getByRole("radio", { name: "Исключить" }));

    expect(onAction).toHaveBeenCalledTimes(1);
    expect(onAction).toHaveBeenCalledWith("slot-open", "exclude");
  });

  it("field edit reports the slot id and the new text", () => {
    const onReplacement = vi.fn();
    render(<WidePlateSlotSection lines={lines} slots={[openSlot]} onReplacement={onReplacement} />);
    expandAttention();

    fireEvent.change(screen.getByRole("textbox", { name: /Заменяющие позиции/ }), {
      target: { value: "ПБ 70-12-8 1" },
    });

    expect(onReplacement).toHaveBeenCalledTimes(1);
    expect(onReplacement).toHaveBeenCalledWith("slot-open", "ПБ 70-12-8 1");
  });
});
