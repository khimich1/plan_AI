import { act, cleanup, renderHook } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { useRecognizedImagePreview } from "@/features/commercial-offer/hooks/useRecognizedImagePreview";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

beforeEach(() => {
  let n = 0;
  vi.spyOn(URL, "createObjectURL").mockImplementation((file: Blob | MediaSource) => {
    n += 1;
    const name = file instanceof File ? file.name : "blob";
    return `blob:${name}:${n}`;
  });
  vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
});

describe("useRecognizedImagePreview.takePreview", () => {
  it("returns file for promote, clears state, and revokes display URL", () => {
    const { result } = renderHook(() => useRecognizedImagePreview());
    const file = new File(["x"], "shot.png", { type: "image/png" });

    act(() => {
      result.current.setPreviewFromFile(file);
    });

    const displayUrl = result.current.preview?.url;
    expect(displayUrl).toBeTruthy();

    let taken: { url: string; name: string; file: File } | null = null;
    act(() => {
      taken = result.current.takePreview();
    });

    expect(taken?.file).toBe(file);
    expect(taken?.name).toBe("shot.png");
    expect(result.current.preview).toBeNull();
    expect(URL.revokeObjectURL).toHaveBeenCalledWith(displayUrl);
  });

  it("returns null when there is no preview", () => {
    const { result } = renderHook(() => useRecognizedImagePreview());

    let taken: unknown = { url: "x", name: "y" };
    act(() => {
      taken = result.current.takePreview();
    });

    expect(taken).toBeNull();
  });
});
