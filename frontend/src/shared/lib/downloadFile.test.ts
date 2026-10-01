import { afterEach, describe, expect, it, vi } from "vitest";

import { saveBlobAs } from "./downloadFile";

describe("saveBlobAs", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("saves a PDF as octet-stream and revokes the URL after 60s", () => {
    vi.useFakeTimers();
    const createObjectURL = vi.fn(() => "blob:pdf");
    const revokeObjectURL = vi.fn();
    vi.stubGlobal("URL", { createObjectURL, revokeObjectURL });

    const anchor = document.createElement("a");
    const click = vi.spyOn(anchor, "click").mockImplementation(() => undefined);
    vi.spyOn(document, "createElement").mockReturnValue(anchor);

    saveBlobAs(new Blob(["%PDF"], { type: "application/pdf" }), "Схема_2026-10-02.pdf");

    expect(anchor.rel).toBe("");
    expect(anchor.download).toBe("Схема_2026-10-02.pdf");
    expect(click).toHaveBeenCalledOnce();
    const created = createObjectURL.mock.calls[0]?.[0] as Blob;
    expect(created.type).toBe("application/octet-stream");
    expect(revokeObjectURL).not.toHaveBeenCalled();

    vi.advanceTimersByTime(59_999);
    expect(revokeObjectURL).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(revokeObjectURL).toHaveBeenCalledOnce();
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:pdf");
  });

  it("keeps a non-pdf blob type", () => {
    vi.useFakeTimers();
    const createObjectURL = vi.fn(() => "blob:sheet");
    const revokeObjectURL = vi.fn();
    vi.stubGlobal("URL", { createObjectURL, revokeObjectURL });

    const anchor = document.createElement("a");
    vi.spyOn(anchor, "click").mockImplementation(() => undefined);
    vi.spyOn(document, "createElement").mockReturnValue(anchor);

    const sheet = new Blob(["xlsx"], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveBlobAs(sheet, "Разбивка.xlsx");

    expect(createObjectURL).toHaveBeenCalledWith(sheet);
    expect(anchor.rel).toBe("");
    vi.advanceTimersByTime(60_000);
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:sheet");
  });
});
