import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";

import { OcrWaitBanner } from "@/features/commercial-offer/components/OcrWaitBanner";
import { OCR_WAIT_MESSAGE } from "@/features/commercial-offer/lib/multiPageSource";

afterEach(() => {
  cleanup();
});

describe("OcrWaitBanner", () => {
  it("shows wait copy and a status region for the spinner", () => {
    render(<OcrWaitBanner />);
    expect(screen.getByRole("status")).toHaveTextContent(OCR_WAIT_MESSAGE);
  });
});
