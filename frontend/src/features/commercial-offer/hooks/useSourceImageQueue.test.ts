import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { useSourceImageQueue } from "@/features/commercial-offer/hooks/useSourceImageQueue";

const makeFile = (name: string) => new File([`content-${name}`], name, { type: "image/png" });

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

describe("useSourceImageQueue", () => {
  it("setFromPages sets length N from page files", () => {
    const createObjectURL = vi.fn((file: File) => `blob:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );

    act(() => {
      result.current.setFromPages([
        { id: "a", file: makeFile("a.png"), name: "a.png" },
        { id: "b", file: makeFile("b.png"), name: "b.png" },
      ]);
    });

    expect(result.current.length).toBe(2);
    expect(result.current.items).toEqual([
      { id: "a", url: "blob:a.png", name: "a.png" },
      { id: "b", url: "blob:b.png", name: "b.png" },
    ]);
    expect(createObjectURL).toHaveBeenCalledTimes(2);
    expect(revokeObjectURL).not.toHaveBeenCalled();
  });

  it("setFromPages again revokes previous urls", () => {
    const createObjectURL = vi.fn((file: File) => `blob:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );

    act(() => {
      result.current.setFromPages([{ id: "a", file: makeFile("a.png"), name: "a.png" }]);
    });
    act(() => {
      result.current.setFromPages([{ id: "b", file: makeFile("b.png"), name: "b.png" }]);
    });

    expect(result.current.length).toBe(1);
    expect(result.current.items[0]?.url).toBe("blob:b.png");
    expect(revokeObjectURL).toHaveBeenCalledTimes(1);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:a.png");
  });

  it("setFromSinglePreview builds one item from File", () => {
    const createObjectURL = vi.fn((file: File) => `blob:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );

    act(() => {
      result.current.setFromSinglePreview(makeFile("solo.png"));
    });

    expect(result.current.length).toBe(1);
    expect(result.current.items[0]).toEqual({
      id: "single",
      url: "blob:solo.png",
      name: "solo.png",
    });
  });

  it("setFromSinglePreview with PreviewLike.file creates independent URL", () => {
    const createObjectURL = vi.fn((file: File) => `blob:fresh:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );
    const file = makeFile("kept.png");

    act(() => {
      result.current.setFromSinglePreview({ url: "blob:stale", name: "kept.png", file });
    });

    expect(result.current.items).toEqual([
      { id: "preview", url: "blob:fresh:kept.png", name: "kept.png" },
    ]);
    expect(createObjectURL).toHaveBeenCalledWith(file);
  });

  it("setFromSinglePreview accepts existing preview and revokes prior queue", () => {
    const createObjectURL = vi.fn((file: File) => `blob:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );

    act(() => {
      result.current.setFromPages([{ id: "a", file: makeFile("a.png"), name: "a.png" }]);
    });
    act(() => {
      result.current.setFromSinglePreview({ url: "blob:kept", name: "kept.png" });
    });

    expect(result.current.items).toEqual([
      { id: "preview", url: "blob:kept", name: "kept.png" },
    ]);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:a.png");
  });

  it("clear empties queue and revokes all urls", () => {
    const createObjectURL = vi.fn((file: File) => `blob:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );

    act(() => {
      result.current.setFromPages([
        { id: "a", file: makeFile("a.png"), name: "a.png" },
        { id: "b", file: makeFile("b.png"), name: "b.png" },
      ]);
    });
    act(() => {
      result.current.clear();
    });

    expect(result.current.length).toBe(0);
    expect(result.current.items).toEqual([]);
    expect(revokeObjectURL).toHaveBeenCalledTimes(2);
  });

  it("unmount revokes leftover urls", () => {
    const createObjectURL = vi.fn((file: File) => `blob:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result, unmount } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );

    act(() => {
      result.current.setFromPages([{ id: "a", file: makeFile("a.png"), name: "a.png" }]);
    });
    unmount();

    expect(revokeObjectURL).toHaveBeenCalledWith("blob:a.png");
  });

  it("parent rerender with default URL factories does not revoke live queue urls", () => {
    const revokeSpy = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
    vi.spyOn(URL, "createObjectURL").mockImplementation((blob: Blob | MediaSource) => {
      const name = blob instanceof File ? blob.name : "blob";
      return `blob:${name}`;
    });

    const { result, rerender } = renderHook(() => useSourceImageQueue());

    act(() => {
      result.current.setFromSinglePreview(makeFile("clipboard-image.png"));
    });

    const url = result.current.items[0]?.url;
    expect(url).toBe("blob:clipboard-image.png");
    expect(revokeSpy).not.toHaveBeenCalled();

    rerender();
    rerender();

    expect(revokeSpy).not.toHaveBeenCalled();
    expect(result.current.items[0]?.url).toBe(url);
  });

  it("replace does not revoke the newly created url", () => {
    const createObjectURL = vi.fn((file: File) => `blob:${file.name}`);
    const revokeObjectURL = vi.fn();
    const { result } = renderHook(() =>
      useSourceImageQueue({ createObjectURL, revokeObjectURL }),
    );

    act(() => {
      result.current.setFromPages([{ id: "a", file: makeFile("a.png"), name: "a.png" }]);
    });
    act(() => {
      result.current.setFromPages([{ id: "b", file: makeFile("b.png"), name: "b.png" }]);
    });

    expect(revokeObjectURL).toHaveBeenCalledTimes(1);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:a.png");
    expect(revokeObjectURL).not.toHaveBeenCalledWith("blob:b.png");
    expect(result.current.items[0]?.url).toBe("blob:b.png");
  });
});
