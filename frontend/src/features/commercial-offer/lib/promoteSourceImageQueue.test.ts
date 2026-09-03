import { describe, expect, it, vi } from "vitest";

import {
  applyPromoteSourceImageQueue,
  type SourceImageQueueSink,
} from "@/features/commercial-offer/lib/promoteSourceImageQueue";

const makeFile = (name: string) => new File([`content-${name}`], name, { type: "image/png" });

const makeSink = (): SourceImageQueueSink & {
  setFromPages: ReturnType<typeof vi.fn>;
  setFromSinglePreview: ReturnType<typeof vi.fn>;
} => ({
  setFromPages: vi.fn(),
  setFromSinglePreview: vi.fn(),
});

describe("applyPromoteSourceImageQueue", () => {
  it("promotes multi-page files into the queue and returns promoted", () => {
    const sink = makeSink();
    const pages = [
      { id: "a", file: makeFile("a.png"), name: "a.png" },
      { id: "b", file: makeFile("b.png"), name: "b.png" },
    ];

    const result = applyPromoteSourceImageQueue(sink, {
      pages,
      singlePreview: { url: "blob:ignored", name: "ignored.png" },
    });

    expect(result).toEqual({ promoted: true, kind: "pages" });
    expect(sink.setFromPages).toHaveBeenCalledWith(pages);
    expect(sink.setFromSinglePreview).not.toHaveBeenCalled();
  });

  it("promotes single preview when pages are empty", () => {
    const sink = makeSink();
    const preview = { url: "blob:solo", name: "solo.png" };

    const result = applyPromoteSourceImageQueue(sink, {
      pages: [],
      singlePreview: preview,
    });

    expect(result).toEqual({ promoted: true, kind: "single" });
    expect(sink.setFromSinglePreview).toHaveBeenCalledWith(preview);
    expect(sink.setFromPages).not.toHaveBeenCalled();
  });

  it("promotes single File when pages are empty", () => {
    const sink = makeSink();
    const file = makeFile("solo.png");

    const result = applyPromoteSourceImageQueue(sink, {
      pages: [],
      singlePreview: file,
    });

    expect(result).toEqual({ promoted: true, kind: "single" });
    expect(sink.setFromSinglePreview).toHaveBeenCalledWith(file);
  });

  it("does nothing for text-only (empty pages, no preview)", () => {
    const sink = makeSink();

    const result = applyPromoteSourceImageQueue(sink, {
      pages: [],
      singlePreview: null,
    });

    expect(result).toEqual({ promoted: false, kind: "none" });
    expect(sink.setFromPages).not.toHaveBeenCalled();
    expect(sink.setFromSinglePreview).not.toHaveBeenCalled();
  });
});
