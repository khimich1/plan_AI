import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";

import { SourceImageQueueControls } from "@/features/commercial-offer/components/SourceImageQueueControls";

afterEach(() => {
  cleanup();
});

describe("SourceImageQueueControls", () => {
  it("renders nothing interactive when queue is empty", () => {
    render(<SourceImageQueueControls items={[]} />);
    expect(screen.queryByRole("button", { name: /Исходные фото/i })).not.toBeInTheDocument();
  });

  it("shows CTA and opens drawer for peer steps", () => {
    render(
      <SourceImageQueueControls
        items={[
          { id: "a", url: "blob:a", name: "a.png" },
          { id: "b", url: "blob:b", name: "b.png" },
        ]}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Исходные фото (2)" }));
    expect(screen.getByRole("dialog")).toBeInTheDocument();
    expect(screen.getByRole("img", { name: "a.png" })).toBeInTheDocument();
  });
});
