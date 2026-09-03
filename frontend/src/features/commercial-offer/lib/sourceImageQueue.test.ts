import { describe, expect, it, vi } from "vitest";

import {
  buildQueueFromPages,
  buildQueueFromPreview,
  clearQueueItems,
  clampQueueIndex,
  nextIndex,
  prevIndex,
  type SourceImageQueueItem,
} from "@/features/commercial-offer/lib/sourceImageQueue";

const makeFile = (name: string) => new File([`content-${name}`], name, { type: "image/png" });

describe("buildQueueFromPages", () => {
  it("builds N items with urls from the factory", () => {
    const pages = [
      { id: "a", file: makeFile("a.png"), name: "a.png" },
      { id: "b", file: makeFile("b.png"), name: "b.png" },
    ];
    const createUrl = vi.fn((file: File) => `blob:${file.name}`);

    const queue = buildQueueFromPages(pages, createUrl);

    expect(queue).toHaveLength(2);
    expect(queue[0]).toEqual({ id: "a", url: "blob:a.png", name: "a.png" });
    expect(queue[1]).toEqual({ id: "b", url: "blob:b.png", name: "b.png" });
    expect(createUrl).toHaveBeenCalledTimes(2);
    expect(createUrl).toHaveBeenNthCalledWith(1, pages[0]!.file);
    expect(createUrl).toHaveBeenNthCalledWith(2, pages[1]!.file);
  });

  it("returns empty queue for empty pages", () => {
    expect(buildQueueFromPages([], vi.fn())).toEqual([]);
  });
});

describe("buildQueueFromPreview", () => {
  it("builds one item from a File via createUrl", () => {
    const file = makeFile("single.png");
    const createUrl = vi.fn((f: File) => `blob:${f.name}`);

    const queue = buildQueueFromPreview(file, createUrl);

    expect(queue).toEqual([{ id: "single", url: "blob:single.png", name: "single.png" }]);
    expect(createUrl).toHaveBeenCalledWith(file);
  });

  it("builds one item from PreviewLike.file via createUrl (independent URL)", () => {
    const file = makeFile("preview.png");
    const createUrl = vi.fn((f: File) => `blob:fresh:${f.name}`);

    const queue = buildQueueFromPreview({ url: "blob:stale", name: "preview.png", file }, createUrl);

    expect(queue).toEqual([{ id: "preview", url: "blob:fresh:preview.png", name: "preview.png" }]);
    expect(createUrl).toHaveBeenCalledWith(file);
  });

  it("builds one item from an existing preview without createUrl", () => {
    const queue = buildQueueFromPreview({ url: "blob:preview", name: "preview.png" });

    expect(queue).toEqual([{ id: "preview", url: "blob:preview", name: "preview.png" }]);
  });

  it("returns empty when preview is nullish", () => {
    expect(buildQueueFromPreview(null)).toEqual([]);
    expect(buildQueueFromPreview(undefined)).toEqual([]);
  });
});

describe("clearQueueItems", () => {
  it("revokes each url exactly once and returns empty", () => {
    const items: SourceImageQueueItem[] = [
      { id: "a", url: "blob:a", name: "a.png" },
      { id: "b", url: "blob:b", name: "b.png" },
    ];
    const revoke = vi.fn();

    const cleared = clearQueueItems(items, revoke);

    expect(cleared).toEqual([]);
    expect(revoke).toHaveBeenCalledTimes(2);
    expect(revoke).toHaveBeenCalledWith("blob:a");
    expect(revoke).toHaveBeenCalledWith("blob:b");
  });

  it("is a no-op for empty items", () => {
    const revoke = vi.fn();
    expect(clearQueueItems([], revoke)).toEqual([]);
    expect(revoke).not.toHaveBeenCalled();
  });
});

describe("pager helpers (clamp, no wrap)", () => {
  it("clampQueueIndex stays within 0..N-1", () => {
    expect(clampQueueIndex(0, 3)).toBe(0);
    expect(clampQueueIndex(2, 3)).toBe(2);
    expect(clampQueueIndex(-1, 3)).toBe(0);
    expect(clampQueueIndex(5, 3)).toBe(2);
    expect(clampQueueIndex(0, 0)).toBe(0);
    expect(clampQueueIndex(1, 1)).toBe(0);
  });

  it("nextIndex / prevIndex clamp without wrap for N=1", () => {
    expect(nextIndex(0, 1)).toBe(0);
    expect(prevIndex(0, 1)).toBe(0);
  });

  it("nextIndex / prevIndex clamp without wrap for N=2", () => {
    expect(nextIndex(0, 2)).toBe(1);
    expect(nextIndex(1, 2)).toBe(1);
    expect(prevIndex(1, 2)).toBe(0);
    expect(prevIndex(0, 2)).toBe(0);
  });
});
